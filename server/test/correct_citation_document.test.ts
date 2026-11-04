import { describe, expect, it } from "vitest";
import request from "supertest";
import { createApp } from "../src/app.js";
import { DocxPackage } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";
import { descendantsW } from "../src/docx/xml.js";
import { correctCitationDocument } from "../src/apa/citations/correct_document.js";
import { correctCitationText } from "../src/apa/citations/correct_text.js";
import { buildDocx } from "./util/docx_builder.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const spec = () => ({ paragraphs: [
  { rawXml: `<w:p><w:pPr><w:jc w:val="right"/><w:ind w:left="240"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Research (Smith and </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Patel 2024) agrees. As Hillebrand et al. (2025: 34) put it, leaders need new forms of management.</w:t></w:r></w:p>` },
  { text: "The National Institute of Standards and Technology (NIST) advises leaders to identify and measure risks across their operations." },
  { text: "An illustration and a link.", image: true, hyperlink: { text: "Source page", url: "https://example.org" } },
  { text: "References", bold: true },
  { text: "Smith, J., & Patel, R. (2024). Leadership. Journal of Leadership, 2, 1–10." },
  { text: "Hillebrand, A., Jones, B., & Lee, C. (2025). New management. Management Journal, 3, 30–40." },
  { text: "National Institute of Standards and Technology. (2023). AI risk management framework. https://example.org/nist" },
], table: { rows: 2, cols: 2, cellText: "A small sample (Smith and Patel 2024) supports the results reported in this table." }, tableAfter: 1 });

describe("citation-only Word correction", () => {
  it("corrects cross-run citations and reference-backed dates while preserving other package parts and formatting", async () => {
    const input = await buildDocx(spec());
    const source = await DocxPackage.load(input);
    const original = await buildDocumentModel(source);
    const result = await correctCitationDocument(input);
    const output = await DocxPackage.load(result.output);
    const model = await buildDocumentModel(output);
    expect(model.paragraphs[0]!.text).toBe("Research (Smith & Patel, 2024) agrees. As Hillebrand et al. (2025, p. 34) put it, leaders need new forms of management.");
    expect(model.paragraphs[1]!.text).toContain("(NIST, 2023)");
    expect(model.paragraphs.filter((p) => p.insideTable).every((p) => !p.text.includes("and Patel 2024"))).toBe(true);
    expect(model.paragraphs.slice(-4).map((p) => p.text)).toEqual(original.paragraphs.slice(-4).map((p) => p.text));
    expect(descendantsW(model.body, "rPr").map((r) => r.toString())).toEqual(descendantsW(original.body, "rPr").map((r) => r.toString()));
    expect(descendantsW(model.body, "pPr").map((r) => r.toString())).toEqual(descendantsW(original.body, "pPr").map((r) => r.toString()));
    expect(output.partNames()).toEqual(source.partNames());
    for (const part of source.partNames()) {
      if (part === "word/document.xml") continue;
      expect(await output.readBinary(part)).toEqual(await source.readBinary(part));
    }
    expect((await correctCitationDocument(result.output)).summary.paragraphsCorrected).toBe(0);
    expect(result.summary.referencesFound).toBe(3);
    expect(result.summary.uncitedReferences).toHaveLength(0);
  });
  it("retains fields and reports missing references", async () => {
    const input = await buildDocx({ paragraphs: [
      { rawXml: `<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>CITATION Smith</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>Smith and Patel 2024</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>` },
      { text: "Some extensive findings were presented in a research paper (Jones, 2022) about the effects of learning and memory in students." },
    ] });
    const result = await correctCitationDocument(input);
    expect(result.output).toEqual(input);
    expect(result.summary.warnings.some((w) => w.includes("fields"))).toBe(true);
    expect(result.summary.warnings.some((w) => w.includes("No reference"))).toBe(true);
  });
  it("uses references supplied with pasted text and keeps the reference list unchanged", () => {
    const text = "National Institute of Standards and Technology (NIST) advises this.\nReferences\nNational Institute of Standards and Technology. (2023). AI risk management framework.";
    expect(correctCitationText(text).text).toBe(text.replace("(NIST)", "(NIST, 2023)"));
  });
  it("uploads and downloads a corrected Word document and rejects unsupported files", async () => {
    const app = createApp();
    const response = await request(app).post("/api/citations/documents").attach("file", await buildDocx(spec()), "paper.docx");
    expect(response.status).toBe(200);
    expect(response.body.paragraphsCorrected).toBeGreaterThan(0);
    const download = await request(app).get(response.body.downloadUrl).buffer(true).parse((res, cb) => {
      const chunks: Buffer[] = []; res.on("data", (c: Buffer) => chunks.push(c));
      res.on("end", () => cb(null, Buffer.concat(chunks)));
    });
    expect(download.status).toBe(200);
    expect(download.headers["content-disposition"]).toContain("citations_corrected.docx");
    const model = await buildDocumentModel(await DocxPackage.load(download.body));
    expect(model.paragraphs[0]!.text).toContain("Smith & Patel, 2024");
    expect((await request(app).post("/api/citations/documents").attach("file", Buffer.from("text"), "paper.pdf")).status).toBe(415);
    expect((await request(app).post("/api/citations/documents")).status).toBe(400);
  });
});
