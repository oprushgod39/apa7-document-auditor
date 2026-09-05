import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import request from "supertest";
import { createApp } from "../src/app.js";
import { getAvailability, getDocument, unlockAi, uploadDocument } from "../src/similarityscan/client.js";

const apiBase = "https://provider.example/functions/v1";
const token = "server-secret-token";

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });
}

beforeEach(() => {
  process.env.SIMILARITYSCAN_API_TOKEN = token;
  process.env.SIMILARITYSCAN_API_BASE_URL = apiBase;
});

afterEach(() => {
  vi.unstubAllGlobals();
  delete process.env.SIMILARITYSCAN_API_TOKEN;
  delete process.env.SIMILARITYSCAN_API_BASE_URL;
});

describe("SimilarityScan documented API contract", () => {
  it("uses the exact availability, upload, status, and AI-unlock endpoints", async () => {
    const calls: Array<{ url: string; method: string; auth: string | null; body?: FormData }> = [];
    vi.stubGlobal("fetch", vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = String(input);
      const headers = new Headers(init?.headers);
      calls.push({ url, method: init?.method ?? "GET", auth: headers.get("authorization"), body: init?.body instanceof FormData ? init.body : undefined });
      if (url.endsWith("/api-status")) return json({ success: true, ready: true, status: "live", acceptingUploads: true });
      if (url.endsWith("/api-documents")) return json({ success: true, orderId: "order-123", status: "submitted" }, 202);
      if (url.includes("/api-document-get?id=order-123")) return json({ success: true, orderId: "order-123", status: "completed", similarity: { percentage: 18 }, ai: { percentage: 42 } });
      if (url.includes("/api-document-unlock-ai?id=order-123")) return json({ success: true, orderId: "order-123", charged: 0.5 });
      throw new Error(`Unexpected URL ${url}`);
    }));

    await getAvailability();
    await uploadDocument({ contents: Buffer.from("A readable test document"), filename: "paper.txt", contentType: "text/plain" });
    await getDocument("order-123");
    await unlockAi("order-123");

    expect(calls.map((call) => call.url)).toEqual([
      `${apiBase}/api-status`,
      `${apiBase}/api-status`,
      `${apiBase}/api-documents`,
      `${apiBase}/api-document-get?id=order-123`,
      `${apiBase}/api-document-unlock-ai?id=order-123`,
    ]);
    expect(calls[0]?.auth).toBeNull();
    expect(calls[2]?.auth).toBe(`Bearer ${token}`);
    expect(calls[2]?.method).toBe("POST");
    expect(calls[2]?.body?.get("checkType")).toBe("full");
    expect(calls[4]?.method).toBe("POST");
  });

  it("maps 429 and 503 responses to safe errors without exposing credentials", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => json({ error: { code: "TOO_MANY_IN_FLIGHT", message: "provider detail" } }, 429)));
    await expect(getAvailability()).rejects.toMatchObject({ httpStatus: 429, code: "NOT_READY" });

    vi.stubGlobal("fetch", vi.fn(async () => json({ error: { code: "SERVICE_UNAVAILABLE", message: "maintenance" } }, 503)));
    await expect(getAvailability()).rejects.toMatchObject({ httpStatus: 503, code: "VERIFICATION_UNAVAILABLE" });
  });
});

describe("/turnitin backend workflow", () => {
  it("uploads, returns real scores, refreshes report URLs, and returns the original", async () => {
    const documentCalls: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (input: string | URL | Request) => {
      const url = String(input);
      if (url.endsWith("/api-status")) return json({ success: true, ready: true, status: "live", acceptingUploads: true });
      if (url.endsWith("/api-documents")) return json({ success: true, orderId: "order-live", status: "submitted" }, 202);
      if (url.includes("/api-document-get?id=order-live")) {
        documentCalls.push(url);
        return json({
          success: true,
          orderId: "order-live",
          fileName: "essay.txt",
          status: "completed",
          similarity: { percentage: 18 },
          ai: { percentage: 42 },
          reports: {
            similarity: { downloadUrl: "https://reports.example/similarity.pdf" },
            ai: { downloadUrl: "https://reports.example/ai.pdf" },
          },
        });
      }
      if (url === "https://reports.example/similarity.pdf") return new Response(Buffer.from("%PDF official similarity"), { headers: { "content-type": "application/pdf" } });
      throw new Error(`Unexpected URL ${url}`);
    }));

    const app = createApp();
    const original = Buffer.from("This is a sufficiently readable English text document for the mocked provider workflow.");
    const uploaded = await request(app).post("/api/similarityscan/documents").attach("file", original, { filename: "essay.txt", contentType: "text/plain" });
    expect(uploaded.status).toBe(202);
    expect(uploaded.body.orderId).toBe("order-live");
    expect(uploaded.body.accessToken).toEqual(expect.any(String));

    const auth = { "X-ScholarlyWorks-Session": uploaded.body.accessToken as string };
    const status = await request(app).get(`/api/similarityscan/documents/${uploaded.body.id}/status`).set(auth);
    expect(status.body).toMatchObject({ status: "completed", similarityPercentage: 18, aiPercentage: 42, similarityReportAvailable: true, aiReportAvailable: true });

    const report = await request(app).get(`/api/similarityscan/documents/${uploaded.body.id}/report/similarity`).set(auth);
    expect(report.status).toBe(200);
    expect(report.headers["content-type"]).toContain("application/pdf");
    expect(documentCalls).toHaveLength(2); // status plus fresh details before download

    const returnedOriginal = await request(app)
      .get(`/api/similarityscan/documents/${uploaded.body.id}/original`)
      .set(auth)
      .buffer(true)
      .parse((res, callback) => {
        const chunks: Buffer[] = [];
        res.on("data", (chunk: Buffer) => chunks.push(chunk));
        res.on("end", () => callback(null, Buffer.concat(chunks)));
      });
    expect(Buffer.compare(returnedOriginal.body as Buffer, original)).toBe(0);

    const unauthorized = await request(app).get(`/api/similarityscan/documents/${uploaded.body.id}/status`);
    expect(unauthorized.status).toBe(404);
    await request(app).delete(`/api/similarityscan/documents/${uploaded.body.id}`).set(auth).expect(204);
  });
});
