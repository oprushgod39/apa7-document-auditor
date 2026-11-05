import { describe, it, expect } from "vitest";
import { parseReference } from "../src/apa/references/parser.js";
import { planEdits, citationLabel } from "../src/apa/citations/evidence/edits.js";
import { publicIPv4 } from "../src/apa/citations/evidence/sources.js";
import { DocxPackage } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";
import { descendantsW } from "../src/docx/xml.js";
import { replaceParagraphTextAt } from "../src/docx/text.js";
import { buildDocx } from "./util/docx_builder.js";
import type { ClaimDecision, EvidenceSource } from "../src/apa/citations/evidence/types.js";

// Synthetic evidence tests the edit gate, not the model's research accuracy.
const source = (raw = "Smith, A. (2024). A study. Journal of Research, 2, 1–10.", id = 0): EvidenceSource => ({ id,
  reference: parseReference(raw, id), url: "https://example.org/study", status: "available", scope: "Synthetic test evidence",
  text: "The study found a sustained increase in children's reading achievement following income support." });
const decision = (sentence = "Income support improved reading achievement."): ClaimDecision => ({ sentence, claim: sentence,
  category: "research", status: "supported", reason: "Supported by the supplied test excerpt.", action: "insert", originalCitation: "",
  anchor: sentence.replace(/\.$/, ""), sourceIds: [0], evidence: [{ sourceId: 0, quote: source().text, identityConfirmed: true }] });

describe("evidence citation edit gate", () => {
  it("finds an uncited claim in a paragraph with another citation and inserts only at its sentence", () => {
    const d = decision(); const text = `Earlier work differs (Jones, 2023). ${d.sentence}`;
    const plan = planEdits(text, 0, [d], [source()]);
    expect(plan.edits).toHaveLength(1);
    expect(text.slice(0, plan.edits[0]!.start) + plan.edits[0]!.value + text.slice(plan.edits[0]!.start)).toBe("Earlier work differs (Jones, 2023). Income support improved reading achievement (Smith, 2024).");
  });
  it("rejects fabricated quotations, title-only evidence, inaccessible sources and metadata conflicts", () => {
    const d = decision();
    for (const s of [{ ...source(), status: "unavailable" as const }, { ...source(), status: "metadata_conflict" as const },
      { ...source(), text: "The title sounds relevant but the source was not retrieved." }]) expect(planEdits(d.sentence, 0, [d], [s]).edits).toHaveLength(0);
    const s = source(); s.reference.title = s.text;
    expect(planEdits(d.sentence, 0, [d], [s]).edits).toHaveLength(0);
  });
  it("does not insert for unsupported statistics, partial support, contradiction, or quotations lacking a verified locator", () => {
    const d = decision("Each dollar yields exactly $1.20 in achievement.");
    for (const status of ["insufficient", "partial", "contradicted"] as const) expect(planEdits(d.sentence, 0, [{ ...d, status }], [source()]).edits).toHaveLength(0);
    expect(planEdits(d.sentence, 0, [{ ...d, category: "quotation" }], [source()]).edits).toHaveLength(0);
  });
  it("replaces an isolated parenthetical citation but preserves narrative prose", () => {
    const text = "Income support improved reading achievement (Wrong, 2007).";
    const d = { ...decision(text), action: "replace" as const, originalCitation: "(Wrong, 2007)", anchor: "" };
    expect(planEdits(text, 0, [d], [source()]).edits[0]?.value).toBe("(Smith, 2024)");
    expect(planEdits("Wrong (2007) explains this.", 0, [{ ...d, sentence: "Wrong (2007) explains this.", claim: "this", originalCitation: "Wrong (2007)" }], [source()]).edits).toHaveLength(0);
  });
  it("formats multiple sources alphabetically, three authors and undated organizations", () => {
    expect(citationLabel(parseReference("Lee, A., Wang, B., & Khan, C. (2024). A study.", 0))).toBe("Lee et al., 2024");
    expect(citationLabel(parseReference("Research Council. (n.d.). Guidance.", 0))).toBe("Research Council, n.d.");
    const d = decision(), a = source("Adams, B. (2021). A study.", 1);
    d.sourceIds.push(1); d.evidence.push({ sourceId: 1, quote: a.text, identityConfirmed: true });
    expect(planEdits(d.sentence, 0, [d], [source(), a]).edits[0]?.value).toBe(" (Adams, 2021; Smith, 2024)");
  });
  it("refuses ambiguous author-year references and duplicate/overlapping proposals", () => {
    const d = decision();
    expect(planEdits(d.sentence, 0, [d], [source(), source(undefined, 1)]).edits).toHaveLength(0);
    expect(planEdits(d.sentence, 0, [d, d], [source()]).edits).toHaveLength(1);
  });
  it("preserves Word runs, tables, images, properties and unrelated package parts after insertion", async () => {
    const d = decision();
    const input = await buildDocx({ paragraphs: [
      { rawXml: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Income support </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>improved reading achievement.</w:t></w:r></w:p>' },
      { text: "Image", image: true, hyperlink: { text: "Source", url: "https://example.org" } },
    ], table: { rows: 2, cols: 2 } });
    const pkg = await DocxPackage.load(input), model = await buildDocumentModel(pkg);
    const properties = descendantsW(model.body, "rPr").map((r) => r.toString());
    const e = planEdits(d.sentence, 0, [d], [source()]).edits[0]!;
    expect(replaceParagraphTextAt(model.paragraphs[0]!.el, e.start, e.length, e.value)).toBe(true);
    pkg.markDirty("word/document.xml");
    const output = await DocxPackage.load(await pkg.save()), out = await buildDocumentModel(output);
    expect(out.paragraphs[0]!.text).toContain("(Smith, 2024)");
    expect(descendantsW(out.body, "rPr").map((r) => r.toString())).toEqual(properties);
    const original = await DocxPackage.load(input);
    for (const part of original.partNames()) if (part !== "word/document.xml") expect(await output.readBinary(part)).toEqual(await original.readBinary(part));
  });
  it("blocks private and special network addresses", () => {
    for (const ip of ["127.0.0.1", "10.1.1.1", "169.254.169.254", "172.20.1.1", "192.168.1.1", "100.64.0.1", "0.0.0.0", "::1"]) expect(publicIPv4(ip)).toBe(false);
    expect(publicIPv4("8.8.8.8")).toBe(true);
  });
});
