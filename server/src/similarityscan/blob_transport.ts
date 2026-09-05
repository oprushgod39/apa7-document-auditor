import { createDecipheriv } from "node:crypto";
import { del, get } from "@vercel/blob";
import { handleUpload, type HandleUploadBody } from "@vercel/blob/client";
import type { Request } from "express";
import { AppError, Errors } from "../errors.js";

export const SIMILARITYSCAN_MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

function requireBlobToken(): string {
  const token = process.env.BLOB_READ_WRITE_TOKEN;
  if (!token) throw new AppError("VERIFICATION_UNAVAILABLE", "Secure large-file upload is not configured.", 503);
  return token;
}

function validBatchId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

function validateBlobUrl(value: string, batchId: string): string {
  let url: URL;
  try { url = new URL(value); } catch { throw Errors.invalid("The secure upload could not be verified."); }
  if (url.protocol !== "https:" || !url.hostname.endsWith(".blob.vercel-storage.com") || !url.pathname.startsWith(`/similarityscan-inputs/${batchId}/`)) {
    throw Errors.invalid("The secure upload could not be verified.");
  }
  return url.toString();
}

function decode(value: string, expectedBytes: number): Buffer {
  if (!/^[A-Za-z0-9_-]+$/.test(value)) throw Errors.invalid("Invalid upload encryption data.");
  const result = Buffer.from(value, "base64url");
  if (result.length !== expectedBytes) throw Errors.invalid("Invalid upload encryption data.");
  return result;
}

export async function authorizeSimilarityScanUpload(req: Request, body: HandleUploadBody): Promise<unknown> {
  return handleUpload({
    request: req,
    body,
    token: requireBlobToken(),
    onBeforeGenerateToken: async (pathname, clientPayload) => {
      let batchId = "";
      try { batchId = String(JSON.parse(clientPayload ?? "{}").batchId ?? ""); } catch { /* rejected below */ }
      if (!validBatchId(batchId) || !pathname.startsWith(`similarityscan-inputs/${batchId}/`) || !pathname.endsWith(".bin")) {
        throw Errors.invalid("Invalid checker upload request.");
      }
      return {
        allowedContentTypes: ["application/octet-stream"],
        maximumSizeInBytes: SIMILARITYSCAN_MAX_UPLOAD_BYTES + 16,
        addRandomSuffix: true,
        tokenPayload: JSON.stringify({ batchId }),
      };
    },
  });
}

export async function readEncryptedUpload(input: {
  url: string;
  batchId: string;
  encryptionKey: string;
  iv: string;
  expectedBytes: number;
}): Promise<Buffer> {
  if (!validBatchId(input.batchId) || input.expectedBytes < 1 || input.expectedBytes > SIMILARITYSCAN_MAX_UPLOAD_BYTES) {
    throw Errors.invalid("Invalid secure upload metadata.");
  }
  const result = await get(validateBlobUrl(input.url, input.batchId), {
    access: "public", token: requireBlobToken(), useCache: false,
  });
  if (!result || result.statusCode !== 200 || result.blob.size > SIMILARITYSCAN_MAX_UPLOAD_BYTES + 16) {
    throw Errors.invalid("The secure upload could not be read.");
  }
  const encrypted = Buffer.from(await new Response(result.stream).arrayBuffer());
  if (encrypted.length < 17 || encrypted.length > SIMILARITYSCAN_MAX_UPLOAD_BYTES + 16) throw Errors.invalid("The secure upload is incomplete.");
  try {
    const decipher = createDecipheriv("aes-256-gcm", decode(input.encryptionKey, 32), decode(input.iv, 12));
    decipher.setAuthTag(encrypted.subarray(encrypted.length - 16));
    const output = Buffer.concat([decipher.update(encrypted.subarray(0, -16)), decipher.final()]);
    if (output.length !== input.expectedBytes) throw new Error("size mismatch");
    return output;
  } catch { throw Errors.invalid("The secure upload could not be decrypted."); }
}

export async function removeEncryptedUpload(url: string, batchId: string): Promise<void> {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return;
  await del(validateBlobUrl(url, batchId), { token: process.env.BLOB_READ_WRITE_TOKEN });
}
