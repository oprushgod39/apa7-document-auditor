import { createCipheriv, createDecipheriv, randomBytes, randomUUID } from "node:crypto";
import { del, get, put } from "@vercel/blob";
import { handleUpload, type HandleUploadBody } from "@vercel/blob/client";
import type { Request } from "express";
import { AppError, Errors } from "../errors.js";
import { config } from "../config.js";

const BINARY_CONTENT_TYPE = "application/octet-stream";
export const MERGE_MAX_TOTAL_BYTES = 200 * 1024 * 1024;

function requireBlobStorage(): string {
  const token = process.env.BLOB_READ_WRITE_TOKEN;
  if (!token) {
    throw new AppError(
      "PROCESSING_FAILED",
      "Large-document merging is temporarily unavailable. Please try again shortly.",
      503
    );
  }
  return token;
}

function validBatchId(batchId: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    batchId
  );
}

function parseBlobUrl(value: string): URL {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw Errors.invalid("An uploaded merger file could not be verified.");
  }
  if (
    url.protocol !== "https:" ||
    !url.hostname.endsWith(".blob.vercel-storage.com")
  ) {
    throw Errors.invalid("An uploaded merger file could not be verified.");
  }
  return url;
}

export function validateMergeBlobUrl(
  value: string,
  batchId: string,
  kind: "merge-inputs" | "merge-outputs"
): string {
  if (!validBatchId(batchId)) throw Errors.invalid("Invalid merger upload batch.");
  const url = parseBlobUrl(value);
  const expectedPrefix = `/${kind}/${batchId}/`;
  if (!url.pathname.startsWith(expectedPrefix)) {
    throw Errors.invalid("An uploaded merger file did not belong to this merge.");
  }
  return url.toString();
}

export async function authorizeMergeUpload(
  req: Request,
  body: HandleUploadBody
): Promise<unknown> {
  const token = requireBlobStorage();
  return handleUpload({
    request: req,
    body,
    token,
    onBeforeGenerateToken: async (pathname, clientPayload) => {
      let batchId = "";
      try {
        batchId = String(JSON.parse(clientPayload ?? "{}").batchId ?? "");
      } catch {
        throw Errors.invalid("Invalid merger upload batch.");
      }
      if (
        !validBatchId(batchId) ||
        !pathname.startsWith(`merge-inputs/${batchId}/`) ||
        !pathname.toLowerCase().endsWith(".bin")
      ) {
        throw Errors.invalid("Only DOCX files can be uploaded to the merger.");
      }
      return {
        allowedContentTypes: [BINARY_CONTENT_TYPE],
        maximumSizeInBytes: config.maxUploadBytes + 32,
        addRandomSuffix: true,
        tokenPayload: JSON.stringify({ batchId }),
      };
    },
  });
}

function decodeSecret(value: string, expectedBytes: number, label: string): Buffer {
  if (!/^[A-Za-z0-9_-]+$/.test(value)) throw Errors.invalid(`Invalid merger ${label}.`);
  const decoded = Buffer.from(value, "base64url");
  if (decoded.length !== expectedBytes) throw Errors.invalid(`Invalid merger ${label}.`);
  return decoded;
}

export function decryptMergePayload(
  encrypted: Buffer,
  encryptionKey: string,
  ivValue: string,
  expectedBytes: number
): Buffer {
  if (encrypted.length < 17) throw Errors.invalid("An encrypted merger file is incomplete.");
  const key = decodeSecret(encryptionKey, 32, "encryption key");
  const iv = decodeSecret(ivValue, 12, "encryption value");
  const authenticationTag = encrypted.subarray(encrypted.length - 16);
  const ciphertext = encrypted.subarray(0, encrypted.length - 16);
  try {
    const decipher = createDecipheriv("aes-256-gcm", key, iv);
    decipher.setAuthTag(authenticationTag);
    const output = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    if (output.length !== expectedBytes || output.length > config.maxUploadBytes) {
      throw Errors.invalid("An uploaded merger file had an unexpected size.");
    }
    return output;
  } catch (error) {
    if (error instanceof AppError) throw error;
    throw Errors.invalid("An uploaded merger file could not be decrypted.");
  }
}

export async function readMergeInput(
  urlValue: string,
  batchId: string,
  encryptionKey: string,
  iv: string,
  expectedBytes: number
): Promise<Buffer> {
  const token = requireBlobStorage();
  const url = validateMergeBlobUrl(urlValue, batchId, "merge-inputs");
  const result = await get(url, { access: "public", token, useCache: false });
  if (!result || result.statusCode !== 200) {
    throw new AppError("PROCESSING_FAILED", "One uploaded document could not be read.", 422);
  }
  if (result.blob.size > config.maxUploadBytes + 32) throw Errors.tooLarge(config.maxUploadBytes);
  const encrypted = Buffer.from(await new Response(result.stream).arrayBuffer());
  if (encrypted.length > config.maxUploadBytes + 32) throw Errors.tooLarge(config.maxUploadBytes);
  return decryptMergePayload(encrypted, encryptionKey, iv, expectedBytes);
}

export async function storeMergeOutput(
  output: Buffer,
  batchId: string,
  encryptionKey: string
) {
  const token = requireBlobStorage();
  if (!validBatchId(batchId)) throw Errors.invalid("Invalid merger upload batch.");
  const key = decodeSecret(encryptionKey, 32, "encryption key");
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(output), cipher.final(), cipher.getAuthTag()]);
  const blob = await put(`merge-outputs/${batchId}/${randomUUID()}.bin`, encrypted, {
    access: "public",
    addRandomSuffix: true,
    contentType: BINARY_CONTENT_TYPE,
    cacheControlMaxAge: 60,
    token,
  });
  return { ...blob, iv: iv.toString("base64url"), size: output.length };
}

export async function cleanupMergeBlobs(
  urls: string[],
  batchId: string,
  kinds: Array<"merge-inputs" | "merge-outputs">
): Promise<void> {
  if (urls.length === 0 || !process.env.BLOB_READ_WRITE_TOKEN) return;
  const allowed = urls.map((value) => {
    for (const kind of kinds) {
      try {
        return validateMergeBlobUrl(value, batchId, kind);
      } catch {
        // Try the next allowed kind.
      }
    }
    throw Errors.invalid("A temporary merger file could not be verified.");
  });
  await del(allowed, { token: process.env.BLOB_READ_WRITE_TOKEN });
}
