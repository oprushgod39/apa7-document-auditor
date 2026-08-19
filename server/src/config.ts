import os from "node:os";
import path from "node:path";

/**
 * Centralized configuration. All values are overridable via environment
 * variables so nothing operational is hard-coded deep in the codebase.
 */
export interface AppConfig {
  productName: string;
  env: string;
  port: number;
  maxUploadBytes: number;
  fileRetentionMinutes: number;
  storageDir: string;
  metadataProvider: "crossref" | "none";
  crossrefBaseUrl: string;
  crossrefMailto: string;
  crossrefTimeoutMs: number;
  crossrefMaxRequestsPerRun: number;
  frontendDistDir: string;
  rateLimitWindowMs: number;
  rateLimitMax: number;
  /** Limits protecting against ZIP bombs / hostile packages. */
  zip: {
    maxEntries: number;
    maxTotalUncompressedBytes: number;
    maxEntryUncompressedBytes: number;
  };
}

function intEnv(name: string, fallback: number): number {
  const v = process.env[name];
  if (!v) return fallback;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export function loadConfig(): AppConfig {
  return {
    // Temporary product name — change here (or via env) to rebrand.
    productName: process.env.PRODUCT_NAME || "APA 7 Document Auditor",
    env: process.env.APP_ENV || process.env.NODE_ENV || "development",
    port: intEnv("PORT", 8000),
    maxUploadBytes: intEnv("MAX_UPLOAD_SIZE", 25 * 1024 * 1024),
    fileRetentionMinutes: intEnv("FILE_RETENTION_MINUTES", 60),
    storageDir:
      process.env.STORAGE_DIR ||
      path.join(os.tmpdir(), "apa7-auditor-storage"),
    metadataProvider:
      (process.env.METADATA_PROVIDER as "crossref" | "none") || "crossref",
    crossrefBaseUrl:
      process.env.CROSSREF_BASE_URL || "https://api.crossref.org",
    crossrefMailto: process.env.CROSSREF_MAILTO || "",
    crossrefTimeoutMs: intEnv("CROSSREF_TIMEOUT_MS", 8000),
    crossrefMaxRequestsPerRun: intEnv("CROSSREF_MAX_REQUESTS_PER_RUN", 25),
    frontendDistDir:
      process.env.FRONTEND_DIST_DIR ||
      path.resolve(import.meta.dirname, "../../web/dist"),
    rateLimitWindowMs: intEnv("RATE_LIMIT_WINDOW_MS", 60_000),
    rateLimitMax: intEnv("RATE_LIMIT_MAX", 60),
    zip: {
      maxEntries: intEnv("ZIP_MAX_ENTRIES", 4096),
      maxTotalUncompressedBytes: intEnv(
        "ZIP_MAX_TOTAL_UNCOMPRESSED",
        400 * 1024 * 1024
      ),
      maxEntryUncompressedBytes: intEnv(
        "ZIP_MAX_ENTRY_UNCOMPRESSED",
        200 * 1024 * 1024
      ),
    },
  };
}

export const config = loadConfig();
