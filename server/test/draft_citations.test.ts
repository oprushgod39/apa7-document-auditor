import { describe, expect, it } from "vitest";
import { buildDocx, malformedStudentPaper } from "./util/docx_builder.js";
import { DocxPackage } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";
import { insertDraftCitations } from "../src/apa/citations/drafts.js";
import { runEngine } from "../src/apa/engine.js";
import { defaultSettings } from "../src/apa/requirements.js";
import { parseCitations } from "../src/apa/citations/parser.js";
import { parseReference } from "../src/apa/references/parser.js";
import { matchCitationsToReferences } from "../src/apa/references/matcher.js";
import { analyzeDocument } from "../src/apa/analysis.js";

describe("draft citation placement", () => {
  it("recognizes organizational authors containing commas and conjunctions", () => {
    const ref = parseReference("Office of Planning, Research and Evaluation. (2012). Third grade follow-up to the Head Start Impact Study: Final report.", 1);
    for (const name of ["Office of Planning, Research and Evaluation", "Office of Planning, Research & Evaluation"]) {
      const result = matchCitationsToReferences(parseCitations(`Findings (${name}, 2012).`, 0), [ref]);
      expect(result.referenceUsage[0]!.cited).toBe(true);
    }
  });
  it("keeps additions as draft findings after reupload, without inserting duplicates", async () => {
    const pkg = await DocxPackage.load(await buildDocx(malformedStudentPaper()));
    const model = await buildDocumentModel(pkg);
    expect(insertDraftCitations(model, new Map())).toHaveLength(6);
    const reloaded = await buildDocumentModel(await DocxPackage.load(await pkg.save()));
    expect(insertDraftCitations(reloaded, new Map())).toHaveLength(0);
    const audit = await runEngine(reloaded, defaultSettings(), { fix: false });
    expect(audit.issues.filter((i) => i.ruleId === "APA-CITATION-DRAFT")).toHaveLength(6);
    const analysis = analyzeDocument(reloaded);
    for (const index of [6, 8, 9]) {
      expect(analysis.citations.filter((c) => c.paragraphIndex === index).length).toBeGreaterThanOrEqual(3);
    }
    expect((await pkg.getXml("word/document.xml")).toString()).not.toContain('w:val="yellow"');
  });
  it("reports insufficient distinct sources instead of repeating a source in the same paragraph", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs.pop();
    const pkg = await DocxPackage.load(await buildDocx(spec));
    insertDraftCitations(await buildDocumentModel(pkg), new Map());
    const model = await buildDocumentModel(await DocxPackage.load(await pkg.save()));
    const analysis = analyzeDocument(model);
    expect(analysis.citations.filter((c) => c.paragraphIndex === 8)).toHaveLength(2);
    const audit = await runEngine(model, { ...defaultSettings(), draftCitationPlacement: true }, { fix: false });
    expect(audit.issues.some((i) => i.originalValue === "shortfall:8")).toBe(true);
  });
  it("does not alter a check-only run even with the draft option selected", async () => {
    const model = await buildDocumentModel(await DocxPackage.load(await buildDocx(malformedStudentPaper())));
    const before = model.paragraphs.map((p) => p.text);
    await runEngine(model, { ...defaultSettings(), mode: "check", draftCitationPlacement: true }, { fix: false });
    expect(model.paragraphs.map((p) => p.text)).toEqual(before);
  });
});
