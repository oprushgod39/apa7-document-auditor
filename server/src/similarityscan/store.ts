import { createHash, randomBytes, randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { kv } from "@vercel/kv";
import { config } from "../config.js";
import { Errors } from "../errors.js";
import { sanitizeFilename } from "../store/sessions.js";
import { readEncryptedUpload, removeEncryptedUpload } from "./blob_transport.js";

export interface SimilarityScanSession {
  id: string;
  createdAt: number;
  orderId: string;
  originalName: string;
  originalSize: number;
  originalContentType: string;
  localPath: string | null;
  blob: { url: string; batchId: string; encryptionKey: string; iv: string } | null;
  accessTokenHash: string;
}

const sessions = new Map<string, SimilarityScanSession>();
const KV_PREFIX = "similarityscan:session:";
const KV_CONFIGURED = Boolean(
  (process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN) || process.env.KV_URL
);

/** A refresh-safe check needs persistent session metadata and encrypted Blob storage. */
export function hasSimilarityScanPersistentStorage(): boolean {
  return KV_CONFIGURED && Boolean(process.env.BLOB_READ_WRITE_TOKEN);
}

function hash(value: string): string {
  return createHash("sha256").update(value).digest("base64url");
}

async function persist(session: SimilarityScanSession): Promise<void> {
  sessions.set(session.id, session);
  if (KV_CONFIGURED) {
    await kv.set(`${KV_PREFIX}${session.id}`, session, { ex: config.similarityScanRetentionDays * 24 * 60 * 60 });
  }
}

export async function createSimilarityScanSession(input: {
  orderId: string;
  originalName: string;
  originalSize: number;
  originalContentType: string;
  buffer?: Buffer;
  blob?: { url: string; batchId: string; encryptionKey: string; iv: string };
}): Promise<{ session: SimilarityScanSession; accessToken: string }> {
  const id = randomUUID();
  const accessToken = randomBytes(32).toString("base64url");
  let localPath: string | null = null;
  if (input.buffer) {
    await fs.mkdir(config.storageDir, { recursive: true });
    localPath = path.join(config.storageDir, `${id}-similarityscan-original.bin`);
    await fs.writeFile(localPath, input.buffer);
  }
  const session: SimilarityScanSession = {
    id,
    createdAt: Date.now(),
    orderId: input.orderId,
    originalName: sanitizeFilename(input.originalName),
    originalSize: input.originalSize,
    originalContentType: input.originalContentType || "application/octet-stream",
    localPath,
    blob: input.blob ?? null,
    accessTokenHash: hash(accessToken),
  };
  await persist(session);
  return { session, accessToken };
}

export async function getSimilarityScanSession(id: string): Promise<SimilarityScanSession> {
  const cached = sessions.get(id);
  if (cached) return cached;
  if (KV_CONFIGURED) {
    const found = await kv.get<SimilarityScanSession>(`${KV_PREFIX}${id}`);
    if (found) { sessions.set(id, found); return found; }
  }
  throw Errors.notFound();
}

export function verifySimilarityScanAccess(session: SimilarityScanSession, token: string | undefined): void {
  if (!token || hash(token) !== session.accessTokenHash) throw Errors.notFound();
}

export async function readSimilarityScanOriginal(session: SimilarityScanSession): Promise<Buffer> {
  if (session.localPath) {
    try { return await fs.readFile(session.localPath); } catch { throw Errors.notFound(); }
  }
  if (session.blob) return readEncryptedUpload({ ...session.blob, expectedBytes: session.originalSize });
  throw Errors.notFound();
}

export async function deleteSimilarityScanSession(id: string): Promise<void> {
  const session = sessions.get(id) ?? (KV_CONFIGURED ? await kv.get<SimilarityScanSession>(`${KV_PREFIX}${id}`) : null);
  if (!session) return;
  sessions.delete(id);
  if (KV_CONFIGURED) await kv.del(`${KV_PREFIX}${id}`);
  if (session.localPath) { try { await fs.unlink(session.localPath); } catch { /* already gone */ } }
  if (session.blob) { try { await removeEncryptedUpload(session.blob.url, session.blob.batchId); } catch { /* best effort */ } }
}
