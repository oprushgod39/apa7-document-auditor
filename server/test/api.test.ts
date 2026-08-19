import { describe, expect, it } from "vitest";
import request from "supertest";
import { createApp } from "../src/app.js";
import { buildDocx, malformedStudentPaper } from "./util/docx_builder.js";
import { DocxPackage, verifyDocxIntegrity } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";

const app = createApp();

async function waitReady(id: string): Promise<void> {
  for (let i = 0; i < 100; i++) {
    const res = await request(app).get(`/api/documents/${id}/status`);
    if (res.body.status === "ready") return;
    if (res.body.status === "error") {
      throw new Error(`processing error: ${res.body.error}`);
    }
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error("timed out waiting for processing");
}

describe("API workflow", () => {
  it("upload → process → report → resolve → download works end to end", async () => {
    const buf = await buildDocx(malformedStudentPaper());

    // 1. Upload
    const up = await request(app)
      .post("/api/documents")
      .attach("file", buf, {
        filename: "Research_Paper.docx",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
    expect(up.status).toBe(201);
    const id = up.body.id as string;
    expect(up.body.detected.references).toBe(3);
    expect(up.body.detected.metadata.title).toContain("Sleep");

    // 2. Process (verifyMetadata off to avoid network in tests)
    const proc = await request(app)
      .post(`/api/documents/${id}/process`)
      .send({ mode: "format_verify", paperType: "student", verifyMetadata: false });
    expect(proc.status).toBe(202);
    await waitReady(id);

    // 3. Report
    const rep = await request(app).get(`/api/documents/${id}/report`);
    expect(rep.status).toBe(200);
    const report = rep.body.report;
    expect(report.state).toBe("review_required");
    expect(report.unresolvedCount).toBeGreaterThan(0);
    expect(report.changesApplied).toBeGreaterThan(5);
    const missing = report.issues.find(
      (i: { ruleId: string }) => i.ruleId === "APA-CITATION-008"
    );
    expect(missing).toBeDefined();

    // 4. Resolve every user-review issue
    for (const issue of report.issues) {
      if (!issue.userResolutionRequired || issue.resolution) continue;
      const key = [
        issue.ruleId,
        issue.location?.paragraphIndex ?? issue.location?.tableIndex ?? "",
        issue.originalValue ?? issue.message,
      ].join("|");
      const optionId = issue.resolutionOptions?.[0]?.id ?? "acknowledge";
      const res = await request(app)
        .post(`/api/documents/${id}/resolve`)
        .send({ issueKey: key, optionId });
      expect(res.status).toBe(200);
    }
    const rep2 = await request(app).get(`/api/documents/${id}/report`);
    expect(rep2.body.report.state).toBe("apa_validated");
    expect(rep2.body.report.unresolvedCount).toBe(0);

    // 5. Download corrected docx — must open without corruption
    const dl = await request(app)
      .get(`/api/documents/${id}/download`)
      .buffer(true)
      .parse((res, cb) => {
        const chunks: Buffer[] = [];
        res.on("data", (c: Buffer) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });
    expect(dl.status).toBe(200);
    expect(dl.headers["content-disposition"]).toContain("Research_Paper_APA7_verified.docx");
    const outBuf = dl.body as Buffer;
    await verifyDocxIntegrity(outBuf);
    const model = await buildDocumentModel(await DocxPackage.load(outBuf));
    expect(model.paragraphs.length).toBeGreaterThan(5);

    // 6. Original is untouched
    const orig = await request(app)
      .get(`/api/documents/${id}/original`)
      .buffer(true)
      .parse((res, cb) => {
        const chunks: Buffer[] = [];
        res.on("data", (c: Buffer) => chunks.push(c));
        res.on("end", () => cb(null, Buffer.concat(chunks)));
      });
    expect(Buffer.compare(orig.body as Buffer, buf)).toBe(0);

    // 7. HTML report
    const html = await request(app).get(`/api/documents/${id}/report.html`);
    expect(html.status).toBe(200);
    expect(html.text).toContain("APA 7 Compliance Report");
  }, 30_000);

  it("rejects unsupported file types cleanly", async () => {
    const res = await request(app)
      .post("/api/documents")
      .attach("file", Buffer.from("%PDF-1.4 fake"), {
        filename: "paper.pdf",
        contentType: "application/pdf",
      });
    expect(res.status).toBe(415);
    expect(res.body.error.code).toBe("UNSUPPORTED_FILE_TYPE");
    expect(res.body.error.message).toContain(".docx");
  });

  it("rejects corrupt docx uploads", async () => {
    const res = await request(app)
      .post("/api/documents")
      .attach("file", Buffer.from("PK\x03\x04 garbage that is not a zip"), {
        filename: "broken.docx",
        contentType: "application/octet-stream",
      });
    expect(res.status).toBe(422);
    expect(res.body.error.code).toBe("CORRUPT_DOCUMENT");
  });

  it("404s for unknown documents without stack traces", async () => {
    const res = await request(app).get("/api/documents/does-not-exist/status");
    expect(res.status).toBe(404);
    expect(res.body.error.code).toBe("DOCUMENT_NOT_FOUND");
    expect(JSON.stringify(res.body)).not.toContain("at ");
  });

  it("applies instructor overrides (font + min references)", async () => {
    const buf = await buildDocx(malformedStudentPaper());
    const up = await request(app)
      .post("/api/documents")
      .attach("file", buf, { filename: "p.docx", contentType: "application/octet-stream" });
    const id = up.body.id as string;
    await request(app)
      .post(`/api/documents/${id}/process`)
      .send({
        mode: "format",
        verifyMetadata: false,
        instructorRequirements: "Times New Roman 12 only\nMinimum five references",
      });
    await waitReady(id);
    const rep = await request(app).get(`/api/documents/${id}/report`);
    expect(rep.body.report.instructorOverrides.join(" ")).toContain("minimum 5 references");
    const minRefIssue = rep.body.report.issues.find(
      (i: { ruleId: string }) => i.ruleId === "APA-REFERENCE-006"
    );
    expect(minRefIssue).toBeDefined(); // only 3 references present
  }, 30_000);
});
