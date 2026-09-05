import { upload } from "@vercel/blob/client";

export interface CheckerAvailability {
  configured: boolean;
  ready: boolean;
  acceptingUploads: boolean;
  /** Vercel KV + Blob are connected, so a refresh can safely resume a check. */
  persistenceEnabled: boolean;
  status: string;
  message?: string;
  retryAfterSeconds?: number;
  avgProcessingMinutes?: number;
  maxUploadBytes: number;
}

export interface CheckerResult {
  id: string;
  orderId: string;
  filename: string;
  status: string;
  terminal: boolean;
  progress: number;
  similarityPercentage: number | null;
  aiPercentage: number | null;
  similarityReportAvailable: boolean;
  aiReportAvailable: boolean;
  aiLocked: boolean;
  aiUnlockPrice: number | null;
  message: string | null;
  accessToken?: string;
}

async function parse<T>(response: Response): Promise<T> {
  if (response.ok) return await response.json() as T;
  let message = response.status === 429
    ? "All checker slots are busy. Please wait and try again."
    : response.status === 503 ? "The checking service is temporarily unavailable." : "The request could not be completed.";
  try { message = (await response.json())?.error?.message ?? message; } catch { /* non-JSON */ }
  throw new Error(message);
}

function sessionHeaders(accessToken: string): HeadersInit {
  return { "X-ScholarlyWorks-Session": accessToken };
}

function base64Url(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

export async function checkerAvailability(): Promise<CheckerAvailability> {
  return parse(await fetch("/api/similarityscan/availability"));
}

export async function createCheck(file: File, useDurableStorage: boolean): Promise<CheckerResult> {
  const local = ["localhost", "127.0.0.1", "::1"].includes(window.location.hostname);
  // Production checks go through encrypted Blob storage when available. This
  // preserves both the original upload and server-side session across refreshes.
  if (local || !useDurableStorage) {
    const form = new FormData();
    form.append("file", file);
    return parse(await fetch("/api/similarityscan/documents", { method: "POST", body: form }));
  }

  const batchId = crypto.randomUUID();
  const key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt"]);
  const rawKey = new Uint8Array(await crypto.subtle.exportKey("raw", key));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, await file.arrayBuffer());
  const stored = await upload(`similarityscan-inputs/${batchId}/source.bin`, new Blob([encrypted]), {
    access: "public",
    handleUploadUrl: "/api/similarityscan/upload",
    clientPayload: JSON.stringify({ batchId }),
    contentType: "application/octet-stream",
    multipart: encrypted.byteLength >= 5 * 1024 * 1024,
  });
  return parse(await fetch("/api/similarityscan/documents", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      batchId,
      url: stored.url,
      encryptionKey: base64Url(rawKey),
      iv: base64Url(iv),
      filename: file.name,
      size: file.size,
      contentType: file.type || "application/octet-stream",
    }),
  }));
}

export async function checkStatus(id: string, accessToken: string): Promise<CheckerResult> {
  return parse(await fetch(`/api/similarityscan/documents/${id}/status`, { headers: sessionHeaders(accessToken) }));
}

export async function unlockAiReport(id: string, accessToken: string): Promise<CheckerResult> {
  return parse(await fetch(`/api/similarityscan/documents/${id}/unlock-ai`, {
    method: "POST", headers: sessionHeaders(accessToken),
  }));
}

export async function viewCheckerReport(id: string, accessToken: string, kind: "similarity" | "ai"): Promise<void> {
  // Open synchronously from the click so browsers do not treat the eventual
  // report view as an unsolicited popup while the backend refreshes its URL.
  const popup = window.open("", "_blank", "noopener,noreferrer");
  try {
    const blob = await reportBlob(id, accessToken, kind);
    const url = URL.createObjectURL(blob);
    if (popup) popup.location.href = url;
    else {
      URL.revokeObjectURL(url);
      throw new Error("Allow pop-ups to view the official report.");
    }
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  } catch (error) {
    popup?.close();
    throw error;
  }
}

export async function downloadCheckerReport(id: string, accessToken: string, kind: "similarity" | "ai"): Promise<void> {
  const blob = await reportBlob(id, accessToken, kind);
  saveBlob(blob, kind === "ai" ? "AI_Report.pdf" : "Similarity_Report.pdf");
}

export async function downloadCheckerOriginal(id: string, accessToken: string, filename: string): Promise<void> {
  const response = await fetch(`/api/similarityscan/documents/${id}/original`, { headers: sessionHeaders(accessToken) });
  if (!response.ok) await parse(response);
  saveBlob(await response.blob(), filename);
}

async function reportBlob(id: string, accessToken: string, kind: "similarity" | "ai"): Promise<Blob> {
  // The backend re-fetches document details on every call, so expired
  // ~7-day provider links are never stored or reused by the browser.
  const response = await fetch(`/api/similarityscan/documents/${id}/report/${kind}`, {
    headers: sessionHeaders(accessToken),
  });
  if (!response.ok) await parse(response);
  return response.blob();
}

function saveBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
