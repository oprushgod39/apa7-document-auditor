import path from "node:path";
import { AppError } from "../errors.js";

const STATUS_PATH = "api-status";
const DOCUMENTS_PATH = "api-documents";
const DOCUMENT_PATH = "api-document-get";
const UNLOCK_AI_PATH = "api-document-unlock-ai";

type ApiErrorBody = { error?: { code?: string; message?: string } | string; code?: string; message?: string };

export interface AvailabilityResponse {
  success: boolean;
  ready: boolean;
  status: "live" | "busy" | "unavailable" | "maintenance" | string;
  acceptingUploads: boolean;
  label?: string;
  message?: string;
  retryAfterSeconds?: number;
  avgProcessingMinutes?: number;
  checkedAt?: string;
}

export interface SimilarityScanDocument {
  success: boolean;
  orderId: string;
  status: "submitted" | "queued" | "processing" | "completed" | "error" | "failed_invalid" | "cancelled" | string;
  fileName?: string;
  similarity?: { percentage?: number };
  ai?: { percentage?: number; locked?: boolean; reason?: string; unlockPrice?: number; unlockUrl?: string };
  reports?: {
    similarity?: { downloadUrl?: string; expiresAt?: string };
    ai?: { downloadUrl?: string; expiresAt?: string };
  };
  message?: string;
  error?: { code?: string; message?: string } | string;
  submittedAt?: string;
  completedAt?: string;
}

export interface UploadResponse {
  success: boolean;
  orderId: string;
  status: string;
  creditsRemaining?: number;
}

function configuration(): { baseUrl: string; token: string } {
  const token = process.env.SIMILARITYSCAN_API_TOKEN?.trim();
  const rawUrl = process.env.SIMILARITYSCAN_API_BASE_URL?.trim();
  if (!token || !rawUrl) {
    throw new AppError(
      "VERIFICATION_UNAVAILABLE",
      "The Turnitin Checker is not connected yet. Add the SimilarityScan credentials to the server environment.",
      503
    );
  }
  let url: URL;
  try { url = new URL(rawUrl); } catch {
    throw new AppError("VERIFICATION_UNAVAILABLE", "SIMILARITYSCAN_API_BASE_URL is invalid.", 503);
  }
  if (url.protocol !== "https:" && process.env.NODE_ENV === "production") {
    throw new AppError("VERIFICATION_UNAVAILABLE", "SIMILARITYSCAN_API_BASE_URL must use HTTPS.", 503);
  }
  url.pathname = url.pathname.replace(/\/+$/, "");
  url.search = "";
  url.hash = "";
  return { baseUrl: url.toString().replace(/\/$/, ""), token };
}

export function isSimilarityScanConfigured(): boolean {
  return Boolean(process.env.SIMILARITYSCAN_API_TOKEN?.trim() && process.env.SIMILARITYSCAN_API_BASE_URL?.trim());
}

function endpoint(baseUrl: string, route: string, orderId?: string): string {
  const url = new URL(`${baseUrl}/${route}`);
  if (orderId) url.searchParams.set("id", orderId);
  return url.toString();
}

function errorDetails(body: unknown): { code: string; message: string } {
  if (!body || typeof body !== "object") return { code: "API_ERROR", message: "" };
  const value = body as ApiErrorBody;
  const nested = value.error && typeof value.error === "object" ? value.error : undefined;
  return {
    code: nested?.code ?? value.code ?? "API_ERROR",
    message: nested?.message ?? value.message ?? (typeof value.error === "string" ? value.error : ""),
  };
}

function publicError(status: number, code: string, message: string): AppError {
  if (status === 429) return new AppError("NOT_READY", "All checker slots are currently busy. Please wait and try again.", 429);
  if (status === 503) return new AppError("VERIFICATION_UNAVAILABLE", message || "The checking service is temporarily unavailable.", 503);
  if (status === 401 || status === 403) return new AppError("VERIFICATION_UNAVAILABLE", "The checker credentials or account access were rejected.", 503);
  if (status === 402) return new AppError("PROCESSING_FAILED", "The SimilarityScan account does not have enough credits for this check.", 402);
  const invalidCodes = new Set(["MISSING_FILE", "FILE_TOO_LARGE", "UNSUPPORTED_TYPE", "NO_READABLE_TEXT", "UNREADABLE_FILE", "UNSUPPORTED_TEXT_VALIDATION"]);
  if (status === 400 || status === 413 || status === 415 || invalidCodes.has(code)) {
    return new AppError(code === "FILE_TOO_LARGE" ? "FILE_TOO_LARGE" : code === "UNSUPPORTED_TYPE" ? "UNSUPPORTED_FILE_TYPE" : "INVALID_REQUEST", message || "This file is not valid for checking.", status);
  }
  return new AppError("PROCESSING_FAILED", message || "The checking service could not complete the request.", status >= 500 ? 503 : 400);
}

async function apiFetch(
  url: string,
  options: RequestInit = {},
  authenticated = true,
  timeoutMs = 30_000
): Promise<Response> {
  const { token } = configuration();
  const headers = new Headers(options.headers);
  if (authenticated) headers.set("Authorization", `Bearer ${token}`);
  let response: Response;
  try {
    response = await fetch(url, { ...options, headers, signal: AbortSignal.timeout(timeoutMs) });
  } catch {
    throw new AppError("VERIFICATION_UNAVAILABLE", "The checking service could not be reached. Please try again.", 503);
  }
  if (!response.ok) {
    let body: unknown = null;
    try { body = await response.clone().json(); } catch { /* non-JSON */ }
    const details = errorDetails(body);
    throw publicError(response.status, details.code, details.message);
  }
  return response;
}

async function apiJson<T>(url: string, options: RequestInit = {}, authenticated = true, timeoutMs?: number): Promise<T> {
  const response = await apiFetch(url, options, authenticated, timeoutMs);
  const body = await response.json() as T;
  if (body && typeof body === "object" && "success" in body && (body as { success?: boolean }).success === false) {
    const details = errorDetails(body);
    throw publicError(400, details.code, details.message);
  }
  return body;
}

export async function getAvailability(): Promise<AvailabilityResponse> {
  const { baseUrl } = configuration();
  return apiJson(endpoint(baseUrl, STATUS_PATH), {}, false, 10_000);
}

export async function uploadDocument(input: {
  contents: Buffer;
  filename: string;
  contentType: string;
}): Promise<UploadResponse> {
  const availability = await getAvailability();
  if (!availability.ready || !availability.acceptingUploads) {
    throw new AppError("VERIFICATION_UNAVAILABLE", availability.message || "The checking service is not accepting uploads right now.", 503);
  }
  const { baseUrl } = configuration();
  const form = new FormData();
  form.append("file", new Blob([input.contents], { type: input.contentType }), path.basename(input.filename));
  form.append("checkType", "full");
  const result = await apiJson<UploadResponse>(endpoint(baseUrl, DOCUMENTS_PATH), { method: "POST", body: form }, true, 55_000);
  if (!result.success || typeof result.orderId !== "string" || !result.orderId) {
    throw new AppError("PROCESSING_FAILED", "The checking service did not return a document ID.", 502);
  }
  return result;
}

export async function getDocument(orderId: string): Promise<SimilarityScanDocument> {
  const { baseUrl } = configuration();
  const result = await apiJson<SimilarityScanDocument>(endpoint(baseUrl, DOCUMENT_PATH, orderId));
  if (!result.success || typeof result.orderId !== "string") {
    throw new AppError("PROCESSING_FAILED", "The checking service returned an invalid document status.", 502);
  }
  return result;
}

export async function unlockAi(orderId: string): Promise<void> {
  const { baseUrl } = configuration();
  await apiJson(endpoint(baseUrl, UNLOCK_AI_PATH, orderId), { method: "POST" });
}

export function validPercentage(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 100 ? value : null;
}

export function reportUrl(document: SimilarityScanDocument, kind: "similarity" | "ai"): string | null {
  const value = document.reports?.[kind]?.downloadUrl;
  if (typeof value !== "string") return null;
  try {
    const url = new URL(value);
    return url.protocol === "https:" ? url.toString() : null;
  } catch { return null; }
}

export async function fetchReport(url: string): Promise<{ contents: Buffer; contentType: string }> {
  let parsed: URL;
  try { parsed = new URL(url); } catch { throw new AppError("VERIFICATION_UNAVAILABLE", "The report URL is invalid.", 502); }
  if (parsed.protocol !== "https:") throw new AppError("VERIFICATION_UNAVAILABLE", "The report URL is invalid.", 502);
  let response: Response;
  try { response = await fetch(parsed, { redirect: "follow", signal: AbortSignal.timeout(30_000) }); }
  catch { throw new AppError("VERIFICATION_UNAVAILABLE", "The report could not be downloaded. Please try again.", 503); }
  if (!response.ok) throw new AppError("VERIFICATION_UNAVAILABLE", "The temporary report link is unavailable. Please try again.", 502);
  const contents = Buffer.from(await response.arrayBuffer());
  if (contents.length === 0 || contents.length > 50 * 1024 * 1024) {
    throw new AppError("VERIFICATION_UNAVAILABLE", "The report download was invalid.", 502);
  }
  return { contents, contentType: response.headers.get("content-type") || "application/pdf" };
}
