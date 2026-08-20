import { randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { kv } from "@vercel/kv";
import { put, del } from "@vercel/blob";
import { config } from "../config.js";
import { Errors } from "../errors.js";
import type { Change, DocumentSettings } from "../apa/types.js";
import type { ComplianceReport } from "../audit/auditor.js";
import type { VerificationResult } from "../verify/provider.js";
import type { DocumentAnalysis } from "../apa/analysis.js";

/**
 * Session store: metadata + document binaries, with two backends.
 *
 * - Local / test fallback (default): in-memory `Map` for metadata and local
 *   temp files on disk under `config.storageDir`. This is the ONLY behavior
 *   exercised by `npm run dev` and the vitest suite, and it is unchanged
 *   from before this file supported a second backend.
 * - Vercel deployment: when `KV_REST_API_URL`/`KV_REST_API_TOKEN` (or
 *   `KV_URL`) are present, session metadata is persisted to Vercel KV so it
 *   survives across separate serverless invocations. When
 *   `BLOB_READ_WRITE_TOKEN` is present, the original/output .docx binaries
 *   are persisted to Vercel Blob instead of local disk (serverless functions
 *   only have ephemeral /tmp, and it is not shared across invocations). The
 *   two are independent: e.g. KV without Blob, or vice versa, both work.
 *
 * Privacy: files live under random, unguessable names/paths and are deleted
 * after FILE_RETENTION_MINUTES. Document contents are never logged. The
 * original upload is preserved untouched so processing can always restart
 * from it.
 */

export type ProcessingStageKey =
  | "read"
  | "structure"
  | "headings"
  | "page_format"
  | "citations"
  | "references"
  | "verify_metadata"
  | "apply"
  | "audit"
  | "prepare_output";

export interface ProcessingStage {
  key: ProcessingStageKey;
  label: string;
  status: "pending" | "running" | "done" | "skipped" | "failed";
}

export const STAGE_LABELS: Record<ProcessingStageKey, string> = {
  read: "Reading Word document",
  structure: "Analyzing document structure",
  headings: "Detecting headings",
  page_format: "Checking page formatting",
  citations: "Analyzing citations",
  references: "Analyzing references",
  verify_metadata: "Verifying scholarly metadata",
  apply: "Applying APA corrections",
  audit: "Running independent APA audit",
  prepare_output: "Preparing corrected Word document",
};

export interface Session {
  id: string;
  createdAt: number;
  /** Sanitized original filename (basename only, safe charset). */
  originalName: string;
  /** Local fs path, or (Blob backend) a Vercel Blob URL. */
  originalPath: string;
  /** Local fs path, or (Blob backend) a Vercel Blob URL. */
  outputPath: string | null;
  settings: DocumentSettings;
  status: "uploaded" | "processing" | "ready" | "error";
  errorMessage: string | null;
  stages: ProcessingStage[];
  report: ComplianceReport | null;
  changes: Change[];
  verification: VerificationResult[] | null;
  /** issueKey → user resolution */
  resolutions: Map<string, { optionId: string; note?: string }>;
  /** Rule IDs the user excluded from formatting. */
  disabledRules: Set<string>;
  /** paragraphIndex → heading level (0 = normal) forced by user resolution. */
  forcedHeadings: Map<number, number>;
  /** Cached analysis of the original document (performance). */
  cachedAnalysis: DocumentAnalysis | null;
}

/** JSON-safe wire form of a Session for KV storage (Map/Set → arrays). */
type SerializedSession = Omit<
  Session,
  "resolutions" | "disabledRules" | "forcedHeadings"
> & {
  resolutions: [string, { optionId: string; note?: string }][];
  disabledRules: string[];
  forcedHeadings: [number, number][];
};

function serialize(s: Session): SerializedSession {
  return {
    ...s,
    resolutions: [...s.resolutions.entries()],
    disabledRules: [...s.disabledRules],
    forcedHeadings: [...s.forcedHeadings.entries()],
  };
}

function deserialize(s: SerializedSession): Session {
  return {
    ...s,
    resolutions: new Map(s.resolutions),
    disabledRules: new Set(s.disabledRules),
    forcedHeadings: new Map(s.forcedHeadings),
  };
}

// --- Backend detection -------------------------------------------------
// Evaluated once at module load. Vercel populates these automatically when
// KV / Blob storage is attached to the project; locally (dev, tests) they
// are absent and the module falls back to in-memory + local fs.
const KV_CONFIGURED = Boolean(
  (process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN) ||
    process.env.KV_URL
);
const BLOB_TOKEN = process.env.BLOB_READ_WRITE_TOKEN;
const BLOB_CONFIGURED = Boolean(BLOB_TOKEN);

function isRemoteRef(p: string): boolean {
  return /^https?:\/\//.test(p);
}

const KV_PREFIX = "apa7:session:";

async function kvGet(id: string): Promise<Session | null> {
  if (!KV_CONFIGURED) return null;
  const raw = await kv.get<SerializedSession>(`${KV_PREFIX}${id}`);
  return raw ? deserialize(raw) : null;
}

async function kvPersist(session: Session): Promise<void> {
  if (!KV_CONFIGURED) return;
  const ttlSeconds = config.fileRetentionMinutes * 60;
  await kv.set(`${KV_PREFIX}${session.id}`, serialize(session), {
    ex: ttlSeconds,
  });
}

async function kvRemove(id: string): Promise<void> {
  if (!KV_CONFIGURED) return;
  await kv.del(`${KV_PREFIX}${id}`);
}

// In-memory map. This is the sole store in local-fallback mode. When KV is
// configured it doubles as a warm-invocation read cache (populated on
// createSession/getSession/saveSession) so repeated access within the same
// serverless container avoids a round trip; KV remains the source of truth
// across separate invocations.
const sessions = new Map<string, Session>();
let sweeperStarted = false;

export function sanitizeFilename(name: string): string {
  const base = path.basename(name).replace(/[^\w.\- ()]/g, "_");
  return base.length > 0 && base !== "." ? base.slice(0, 120) : "document.docx";
}

export async function ensureStorageDir(): Promise<void> {
  await fs.mkdir(config.storageDir, { recursive: true });
}

const DOCX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

async function storeBinary(
  sessionId: string,
  role: "original" | "output",
  buffer: Buffer
): Promise<string> {
  if (BLOB_CONFIGURED) {
    // Note: Vercel Blob has no native TTL. We rely on `deleteSession` to
    // remove blobs eagerly (explicit delete, or the retention sweeper in
    // fallback mode). When running on the KV backend, session metadata
    // expires via KV TTL rather than an interval sweep (see startSweeper);
    // if a session's KV record expires without an explicit delete, its
    // blobs are NOT automatically cleaned up (Blob supports no TTL). Given
    // document sizes are small and retention windows short, we accept this
    // as an acceptable trade-off rather than adding a scheduled cleanup
    // job — revisit with a Vercel Cron job hitting a cleanup endpoint if
    // orphaned blobs become a storage concern.
    const blob = await put(`sessions/${sessionId}-${role}.docx`, buffer, {
      access: "public",
      addRandomSuffix: true,
      contentType: DOCX_CONTENT_TYPE,
      token: BLOB_TOKEN,
    });
    return blob.url;
  }
  await ensureStorageDir();
  const filePath = path.join(config.storageDir, `${sessionId}-${role}.docx`);
  await fs.writeFile(filePath, buffer);
  return filePath;
}

export async function createSession(
  originalName: string,
  buffer: Buffer,
  settings: DocumentSettings
): Promise<Session> {
  const id = randomUUID();
  const originalPath = await storeBinary(id, "original", buffer);
  const session: Session = {
    id,
    createdAt: Date.now(),
    originalName: sanitizeFilename(originalName),
    originalPath,
    outputPath: null,
    settings,
    status: "uploaded",
    errorMessage: null,
    stages: [],
    report: null,
    changes: [],
    verification: null,
    resolutions: new Map(),
    disabledRules: new Set(),
    forcedHeadings: new Map(),
    cachedAnalysis: null,
  };
  sessions.set(id, session);
  if (KV_CONFIGURED) {
    await kvPersist(session);
  } else {
    startSweeper();
  }
  return session;
}

export async function getSession(id: string): Promise<Session> {
  const cached = sessions.get(id);
  if (cached) return cached;
  if (KV_CONFIGURED) {
    const fetched = await kvGet(id);
    if (fetched) {
      sessions.set(id, fetched);
      return fetched;
    }
  }
  throw Errors.notFound();
}

/**
 * Persist a mutated Session back to the store. Every call site that
 * mutates a Session field in place (`session.status = ...`, etc.) must call
 * this afterward so the change is visible on the KV backend, where each
 * serverless invocation may be a separate process. In local-fallback mode
 * this is a cheap Map re-set — the in-memory Map already holds the live
 * object reference, so it's effectively a no-op.
 */
export async function saveSession(session: Session): Promise<void> {
  sessions.set(session.id, session);
  if (KV_CONFIGURED) {
    await kvPersist(session);
  }
}

export async function readOriginal(session: Session): Promise<Buffer> {
  try {
    if (isRemoteRef(session.originalPath)) {
      const res = await fetch(session.originalPath);
      if (!res.ok) throw new Error(`blob fetch failed: ${res.status}`);
      return Buffer.from(await res.arrayBuffer());
    }
    return await fs.readFile(session.originalPath);
  } catch {
    throw Errors.notFound();
  }
}

export async function writeOutput(session: Session, buffer: Buffer): Promise<void> {
  session.outputPath = await storeBinary(session.id, "output", buffer);
  await saveSession(session);
}

export async function readOutput(session: Session): Promise<Buffer> {
  if (!session.outputPath) throw Errors.notFound();
  if (isRemoteRef(session.outputPath)) {
    const res = await fetch(session.outputPath);
    if (!res.ok) throw Errors.notFound();
    return Buffer.from(await res.arrayBuffer());
  }
  return await fs.readFile(session.outputPath);
}

export async function deleteSession(id: string): Promise<void> {
  const s = sessions.get(id) ?? (KV_CONFIGURED ? await kvGet(id) : undefined);
  if (!s) return;
  sessions.delete(id);
  if (KV_CONFIGURED) await kvRemove(id);
  for (const p of [s.originalPath, s.outputPath]) {
    if (!p) continue;
    try {
      if (isRemoteRef(p)) {
        if (BLOB_CONFIGURED) await del(p, { token: BLOB_TOKEN });
      } else {
        await fs.unlink(p);
      }
    } catch {
      /* already gone */
    }
  }
}

function startSweeper(): void {
  // Only meaningful in local-fallback mode: the KV backend expires session
  // metadata via TTL instead (see kvPersist), since a manual interval sweep
  // cannot run reliably across independent serverless invocations.
  if (sweeperStarted || KV_CONFIGURED) return;
  sweeperStarted = true;
  const interval = setInterval(async () => {
    const cutoff = Date.now() - config.fileRetentionMinutes * 60_000;
    for (const [id, s] of sessions) {
      if (s.createdAt < cutoff && s.status !== "processing") {
        await deleteSession(id);
      }
    }
  }, 60_000);
  interval.unref();
}

/** For tests. */
export function _clearSessions(): void {
  sessions.clear();
}
