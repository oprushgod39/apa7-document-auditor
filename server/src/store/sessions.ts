import { randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { config } from "../config.js";
import { Errors } from "../errors.js";
import type { Change, DocumentSettings } from "../apa/types.js";
import type { ComplianceReport } from "../audit/auditor.js";
import type { VerificationResult } from "../verify/provider.js";
import type { DocumentAnalysis } from "../apa/analysis.js";

/**
 * Session store: in-memory metadata + temp files on disk.
 *
 * Privacy: files live under a dedicated storage dir with random names and
 * are deleted after FILE_RETENTION_MINUTES. Document contents are never
 * logged. The original upload is preserved untouched so processing can
 * always restart from it.
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
  originalPath: string;
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

const sessions = new Map<string, Session>();
let sweeperStarted = false;

export function sanitizeFilename(name: string): string {
  const base = path.basename(name).replace(/[^\w.\- ()]/g, "_");
  return base.length > 0 && base !== "." ? base.slice(0, 120) : "document.docx";
}

export async function ensureStorageDir(): Promise<void> {
  await fs.mkdir(config.storageDir, { recursive: true });
}

export async function createSession(
  originalName: string,
  buffer: Buffer,
  settings: DocumentSettings
): Promise<Session> {
  await ensureStorageDir();
  const id = randomUUID();
  const originalPath = path.join(config.storageDir, `${id}-original.docx`);
  await fs.writeFile(originalPath, buffer);
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
  startSweeper();
  return session;
}

export function getSession(id: string): Session {
  const s = sessions.get(id);
  if (!s) throw Errors.notFound();
  return s;
}

export async function readOriginal(session: Session): Promise<Buffer> {
  try {
    return await fs.readFile(session.originalPath);
  } catch {
    throw Errors.notFound();
  }
}

export async function writeOutput(session: Session, buffer: Buffer): Promise<void> {
  const outputPath = path.join(config.storageDir, `${session.id}-output.docx`);
  await fs.writeFile(outputPath, buffer);
  session.outputPath = outputPath;
}

export async function deleteSession(id: string): Promise<void> {
  const s = sessions.get(id);
  if (!s) return;
  sessions.delete(id);
  for (const p of [s.originalPath, s.outputPath]) {
    if (!p) continue;
    try {
      await fs.unlink(p);
    } catch {
      /* already gone */
    }
  }
}

function startSweeper(): void {
  if (sweeperStarted) return;
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
