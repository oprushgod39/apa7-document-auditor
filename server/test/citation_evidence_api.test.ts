import { afterEach, describe, expect, it, vi } from "vitest";
import request from "supertest";
import { createApp } from "../src/app.js";
import { buildDocx } from "./util/docx_builder.js";
import { DocxPackage } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";
import { retrieveSource } from "../src/apa/citations/evidence/sources.js";
import { assessParagraph } from "../src/apa/citations/evidence/openai.js";
import type { ClaimDecision } from "../src/apa/citations/evidence/types.js";

vi.mock("../src/apa/citations/evidence/sources.js", () => ({ retrieveSource: vi.fn() }));
vi.mock("../src/apa/citations/evidence/openai.js", () => ({ evidenceConfigured: () => Boolean(process.env.OPENAI_API_KEY && process.env.OPENAI_CITATION_MODEL), assessParagraph: vi.fn() }));
afterEach(() => { vi.unstubAllEnvs(); vi.resetAllMocks(); });
describe("resumable evidence API (simulated provider responses)", () => {
  it("rejects unconfigured verification", async () => {
    vi.stubEnv("OPENAI_API_KEY", ""); vi.stubEnv("OPENAI_CITATION_MODEL", "");
    const app = createApp(); expect((await request(app).get("/api/citations/verification/availability")).body.available).toBe(false);
    await request(app).post("/api/citations/verification").attach("file", await buildDocx({ paragraphs: [{ text: "A document." }] }), "test.docx").expect(503);
  });
  it("resumes failures, requires token, applies a source-grounded edit once, and downloads both artifacts", async () => {
    vi.stubEnv("OPENAI_API_KEY", "test-not-real"); vi.stubEnv("OPENAI_CITATION_MODEL", "test-model");
    const sentence = "Income support improved reading achievement among children in the study.";
    const quote = "The study found an increase in reading achievement among children receiving income support.";
    vi.mocked(retrieveSource).mockImplementation(async (reference, id) => ({ id, reference, url: "https://example.org", text: quote, status: "available", scope: "Synthetic API test evidence" }));
    vi.mocked(assessParagraph).mockRejectedValueOnce(new Error("temporary provider failure")).mockImplementation(async (text) => text === sentence ? [{ sentence, claim: sentence, category: "research", status: "supported", reason: "Synthetic test", evidence: [{ sourceId: 0, quote, identityConfirmed: true }], action: "insert", originalCitation: "", anchor: sentence.slice(0, -1), sourceIds: [0] } satisfies ClaimDecision] : []);
    const app = createApp();
    const input = await buildDocx({ paragraphs: [{ text: sentence, italic: true }, { text: "References", bold: true }, { text: "Smith, A. (2024). Income support. Journal of Research, 2, 1–10." }, { text: "Unused, B. (2020). Other study. Journal of Research, 3, 11–20." }] });
    const start = await request(app).post("/api/citations/verification").attach("file", input, "test.docx").expect(200);
    const { id, token } = start.body; const url = `/api/citations/verification/${id}`;
    await request(app).get(url).expect(404);
    let state = start.body;
    while (state.phase === "sources") state = (await request(app).post(url + "/step").set("X-Citation-Token", token).send({ version: state.version }).expect(200)).body;
    await request(app).post(url + "/step").set("X-Citation-Token", token).send({ version: state.version }).expect(500);
    expect((await request(app).get(url).set("X-Citation-Token", token)).body.version).toBe(state.version);
    for (let i = 0; i < 15 && state.phase !== "complete"; i++) state = (await request(app).post(url + "/step").set("X-Citation-Token", token).send({ version: state.version }).expect(200)).body;
    expect(state.phase).toBe("complete"); expect(state.applied).toBe(1);
    const repeated = await request(app).post(url + "/step").set("X-Citation-Token", token).send({ version: 0 }).expect(200);
    expect(repeated.body.applied).toBe(1);
    const doc = await request(app).get(url + "/document").set("X-Citation-Token", token).buffer(true).parse((res, callback) => { const chunks: Buffer[] = []; res.on("data", (b) => chunks.push(b)); res.on("end", () => callback(null, Buffer.concat(chunks))); }).expect(200);
    const model = await buildDocumentModel(await DocxPackage.load(doc.body));
    expect(model.paragraphs[0]!.text).toBe(sentence.slice(0, -1) + " (Smith, 2024).");
    const report = await request(app).get(url + "/report").set("X-Citation-Token", token).expect(200);
    expect(report.text).toContain(quote); expect(report.text).toContain("Unused, B."); expect(report.text).toContain("correction applied");
  });
});
