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
    expect(insertDraftCitations(model, new Map())).toHaveLength(1);
    const reloaded = await buildDocumentModel(await DocxPackage.load(await pkg.save()));
    expect(insertDraftCitations(reloaded, new Map())).toHaveLength(0);
    const audit = await runEngine(reloaded, defaultSettings(), { fix: false });
    expect(audit.issues.filter((i) => i.ruleId === "APA-CITATION-DRAFT")).toHaveLength(1);
  });
  it("does not alter a check-only run even with the draft option selected", async () => {
    const model = await buildDocumentModel(await DocxPackage.load(await buildDocx(malformedStudentPaper())));
    const before = model.paragraphs.map((p) => p.text);
    await runEngine(model, { ...defaultSettings(), mode: "check", draftCitationPlacement: true }, { fix: false });
    expect(model.paragraphs.map((p) => p.text)).toEqual(before);
  });
});
