import { randomUUID } from "node:crypto";
import { del, get, put } from "@vercel/blob";
import { handleUpload, type HandleUploadBody } from "@vercel/blob/client";
import type { Request } from "express";
import { AppError, Errors } from "../errors.js";
import { config } from "../config.js";

const DOCX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PDF_CONTENT_TYPE = "application/pdf";
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
        !pathname.toLowerCase().endsWith(".docx")
      ) {
        throw Errors.invalid("Only DOCX files can be uploaded to the merger.");
      }
      return {
        allowedContentTypes: [DOCX_CONTENT_TYPE, "application/octet-stream", "application/zip"],
        maximumSizeInBytes: config.maxUploadBytes,
        addRandomSuffix: true,
        tokenPayload: JSON.stringify({ batchId }),
      };
    },
  });
}

export async function readMergeInput(urlValue: string, batchId: string): Promise<Buffer> {
  const token = requireBlobStorage();
  const url = validateMergeBlobUrl(urlValue, batchId, "merge-inputs");
  const result = await get(url, { access: "public", token, useCache: false });
  if (!result || result.statusCode !== 200) {
    throw new AppError("PROCESSING_FAILED", "One uploaded document could not be read.", 422);
  }
  if (result.blob.size > config.maxUploadBytes) throw Errors.tooLarge(config.maxUploadBytes);
  const buffer = Buffer.from(await new Response(result.stream).arrayBuffer());
  if (buffer.length > config.maxUploadBytes) throw Errors.tooLarge(config.maxUploadBytes);
  return buffer;
}

export async function storeMergeOutput(output: Buffer, batchId: string) {
  const token = requireBlobStorage();
  if (!validBatchId(batchId)) throw Errors.invalid("Invalid merger upload batch.");
  return put(`merge-outputs/${batchId}/${randomUUID()}-Merged_Submissions.pdf`, output, {
    access: "public",
    addRandomSuffix: true,
    contentType: PDF_CONTENT_TYPE,
    cacheControlMaxAge: 60,
    token,
  });
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
