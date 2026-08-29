/** Typed API client for the APA 7 Document Auditor backend. */

import { upload } from "@vercel/blob/client";

export interface DetectedInfo {
  metadata: {
    title?: string;
    author?: string;
    institution?: string;
    courseNumber?: string;
    courseName?: string;
    instructor?: string;
    dueDate?: string;
  };
  hasTitlePage: boolean;
  hasAbstract: boolean;
  headings: number;
  citations: number;
  references: number;
  tables: number;
  images: number;
  paragraphs: number;
  footnotes: number;
}

export interface Stage {
  key: string;
  label: string;
  status: "pending" | "running" | "done" | "skipped" | "failed";
}

export interface SessionInfo {
  id: string;
  originalName: string;
  status: "uploaded" | "processing" | "ready" | "error";
  stages: Stage[];
  error: string | null;
  settings: {
    paperType: string;
    mode: string;
    preserveWording: boolean;
    fixCitationMechanics: boolean;
    verifyMetadata: boolean;
  };
}

export interface UploadResponse extends SessionInfo {
  detected: DetectedInfo;
}

export interface ResolutionOption {
  id: string;
  label: string;
  description?: string;
}

export interface Issue {
  id: string;
  ruleId: string;
  category: string;
  severity: "error" | "warning" | "info";
  status: string;
  message: string;
  explanation?: string;
  location?: {
    paragraphIndex?: number;
    tableIndex?: number;
    description?: string;
    excerpt?: string;
  };
  originalValue?: string;
  suggestedValue?: string;
  confidence: number;
  autoFixable: boolean;
  userResolutionRequired: boolean;
  resolutionOptions?: ResolutionOption[];
  resolution?: { optionId: string; note?: string };
  resolved: boolean;
}

export interface Change {
  id: string;
  ruleId: string;
  category: string;
  location: { paragraphIndex?: number; description?: string; excerpt?: string };
  before: string;
  after: string;
  reason: string;
  confidence: number;
  stage: string;
  documentWide?: boolean;
}

export interface CategorySummary {
  category: string;
  label: string;
  status: string;
  rulesChecked: number;
  rulesPassed: number;
  issueCount: number;
}

export interface VerificationSummary {
  provider: string;
  attempted: number;
  verified: number;
  probable: number;
  mismatched: number;
  unverified: number;
  providerUnavailable: boolean;
}

export interface Report {
  state: "apa_validated" | "review_required";
  unresolvedCount: number;
  rulesResolvedPercent: number;
  categories: CategorySummary[];
  issues: Issue[];
  changes: Change[];
  changesApplied: number;
  verification: VerificationSummary | null;
  instructorOverrides: string[];
  generatedAt: string;
}

export interface OutlineEntry {
  index: number;
  kind: string;
  level?: number;
  confidence?: string;
  text: string;
  issues: number;
}

export interface ReportResponse extends SessionInfo {
  report: Report;
  outline: OutlineEntry[];
  instructorUninterpreted: string[];
}

export interface ProcessSettings {
  paperType: "student" | "professional";
  mode: "check" | "format" | "format_verify";
  preserveWording: boolean;
  fixCitationMechanics: boolean;
  verifyMetadata: boolean;
  metadata: Record<string, string>;
  instructorRequirements: string;
}

export class ApiError extends Error {
  code: string;
  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

async function handle<T>(res: Response): Promise<T> {
  if (res.ok) return (await res.json()) as T;
  let code = "INTERNAL";
  let message = "Something went wrong.";
  try {
    const body = await res.json();
    code = body?.error?.code ?? code;
    message = body?.error?.message ?? message;
  } catch {
    /* non-JSON error */
  }
  throw new ApiError(code, message);
}

export async function uploadDocument(file: File): Promise<UploadResponse> {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch("/api/documents", { method: "POST", body: form });
  return handle<UploadResponse>(res);
}

export async function startProcessing(
  id: string,
  settings: ProcessSettings
): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/process`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(settings),
  });
  return handle<SessionInfo>(res);
}

export async function getStatus(id: string): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/status`);
  return handle<SessionInfo>(res);
}

export async function getReport(id: string): Promise<ReportResponse> {
  const res = await fetch(`/api/documents/${id}/report`);
  return handle<ReportResponse>(res);
}

export async function resolveIssue(
  id: string,
  issueKey: string,
  optionId: string
): Promise<{ report: Report }> {
  const res = await fetch(`/api/documents/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ issueKey, optionId }),
  });
  return handle<{ report: Report }>(res);
}

export async function regenerate(id: string): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/generate`, { method: "POST" });
  return handle<SessionInfo>(res);
}

export function issueKeyOf(issue: Issue): string {
  return [
    issue.ruleId,
    issue.location?.paragraphIndex ?? issue.location?.tableIndex ?? "",
    issue.originalValue ?? issue.message,
  ].join("|");
}

export const downloadUrl = (id: string) => `/api/documents/${id}/download`;
export const reportDownloadUrl = (id: string) => `/api/documents/${id}/report.html`;

const MERGE_DIRECT_LIMIT = 3 * 1024 * 1024;
const MERGE_WORD_PATTERN = /[\p{L}\p{N}]+(?:['’.-][\p{L}\p{N}]+)*/gu;
const MERGE_REFERENCE_HEADING = /^(references|bibliography|works\s+cited)\s*:?\s*[.]*$/i;

function countMergeWords(text: string): number {
  return text.match(MERGE_WORD_PATTERN)?.length ?? 0;
}

async function countMergeFile(file: File): Promise<number> {
  const mammoth = (await import("mammoth")).default;
  const output = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
  const lines = (output.value ?? "").replace(/\r\n?/g, "\n").split("\n");
  const referenceIndex = lines.findIndex((line) => MERGE_REFERENCE_HEADING.test(line.trim()));
  return countMergeWords(lines.slice(0, referenceIndex >= 0 ? referenceIndex : undefined).join("\n"));
}

export async function previewMergeDocuments(files: File[]): Promise<{ contentWords: number[]; appendixSourceWords: number }> {
  // Count in the browser so selecting a large batch never sends it through
  // the hosted request-body limit merely to update the real-time display.
  const contentWords: number[] = [];
  for (const file of files) contentWords.push(await countMergeFile(file));
  const info = await fetch("/api/merge-info");
  const { appendixSourceWords } = await handle<{ appendixSourceWords: number }>(info);
  return { contentWords, appendixSourceWords };
}

export async function mergeDocuments(items: { file: File; name: string }[], appendixWords: number): Promise<Blob> {
  const totalBytes = items.reduce((sum, item) => sum + item.file.size, 0);
  const hostname = typeof window === "undefined" ? "" : window.location.hostname;
  const localHost = hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
  if (totalBytes > MERGE_DIRECT_LIMIT && !localHost) {
    return mergeLargeDocuments(items, appendixWords);
  }
  const form = new FormData();
  for (const item of items) form.append("documents", item.file);
  form.append("names", JSON.stringify(items.map((item) => item.name)));
  form.append("appendixWords", String(appendixWords));
  const res = await fetch("/api/merge-documents", { method: "POST", body: form });
  if (res.ok) return res.blob();
  let message = res.status === 413
    ? "This upload is too large for direct processing. Please try again."
    : "The documents could not be merged.";
  try {
    const body = await res.json();
    message = body?.error?.message ?? message;
  } catch { /* non-JSON error */ }
  throw new Error(message);
}

function base64Url(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function fromBase64Url(value: string): Uint8Array<ArrayBuffer> {
  const base64 = value.replace(/-/g, "+").replace(/_/g, "/");
  const binary = atob(base64 + "=".repeat((4 - base64.length % 4) % 4));
  const bytes = new Uint8Array(new ArrayBuffer(binary.length));
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

async function cleanupMergeBatch(batchId: string, urls: string[]): Promise<void> {
  if (urls.length === 0) return;
  try {
    await fetch("/api/merge-cleanup", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId, urls }),
    });
  } catch {
    // Cleanup is best-effort; it must not hide the useful merge result/error.
  }
}

async function mergeLargeDocuments(
  items: { file: File; name: string }[],
  appendixWords: number
): Promise<Blob> {
  const batchId = crypto.randomUUID();
  const encryptionKey = await crypto.subtle.generateKey(
    { name: "AES-GCM", length: 256 },
    true,
    ["encrypt", "decrypt"]
  );
  const exportedKey = new Uint8Array(await crypto.subtle.exportKey("raw", encryptionKey));
  const encryptionKeyValue = base64Url(exportedKey);
  const uploaded: Array<{ url: string; file: File; name: string; iv: string }> = [];
  let outputUrl = "";
  try {
    for (let index = 0; index < items.length; index++) {
      const item = items[index]!;
      const iv = crypto.getRandomValues(new Uint8Array(12));
      const encrypted = await crypto.subtle.encrypt(
        { name: "AES-GCM", iv },
        encryptionKey,
        await item.file.arrayBuffer()
      );
      const blob = await upload(`merge-inputs/${batchId}/${index + 1}.bin`, new Blob([encrypted]), {
        access: "public",
        handleUploadUrl: "/api/merge-upload",
        clientPayload: JSON.stringify({ batchId }),
        contentType: "application/octet-stream",
        multipart: encrypted.byteLength >= 5 * 1024 * 1024,
      });
      uploaded.push({ url: blob.url, file: item.file, name: item.name, iv: base64Url(iv) });
    }

    const mergeResponse = await fetch("/api/merge-documents-from-blobs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        batchId,
        documents: uploaded.map((entry) => ({
          url: entry.url,
          name: entry.name,
          originalName: entry.file.name,
          size: entry.file.size,
          iv: entry.iv,
        })),
        appendixWords,
        encryptionKey: encryptionKeyValue,
      }),
    });
    const result = await handle<{
      url: string;
      downloadUrl: string;
      filename: string;
      iv: string;
      size: number;
    }>(mergeResponse);
    outputUrl = result.url;
    const download = await fetch(result.downloadUrl);
    if (!download.ok) throw new Error("The merged PDF was created but could not be downloaded.");
    const decrypted = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: fromBase64Url(result.iv) },
      encryptionKey,
      await download.arrayBuffer()
    );
    if (decrypted.byteLength !== result.size) {
      throw new Error("The merged PDF download was incomplete.");
    }
    return new Blob([decrypted], { type: "application/pdf" });
  } catch (cause) {
    if (cause instanceof ApiError) throw cause;
    const message = cause instanceof Error ? cause.message : "The documents could not be merged.";
    throw new Error(message.includes("BLOB_READ_WRITE_TOKEN")
      ? "Large-document merging is not configured on the published website yet."
      : message);
  } finally {
    await cleanupMergeBatch(batchId, [
      ...uploaded.map((entry) => entry.url),
      ...(outputUrl ? [outputUrl] : []),
    ]);
  }
}
