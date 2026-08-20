import { describe, expect, it } from "vitest";
import { DocxPackage, verifyDocxIntegrity } from "../src/docx/package.js";
import { buildDocumentModel, contentFingerprint } from "../src/docx/model.js";
import { analyzeDocument } from "../src/apa/analysis.js";
import { runEngine } from "../src/apa/engine.js";
import { assertContentPreserved } from "../src/pipeline.js";
import { defaultSettings } from "../src/apa/requirements.js";
import { buildDocx, type DocSpec } from "./util/docx_builder.js";

/**
 * Reproduces the real-world defect: "References" typed as a bold run at the
 * tail of the final body paragraph, separated only by a manual line break
 * (Shift+Enter, `<w:br/>` with no type) rather than living in its own
 * paragraph — confirmed against a real submitted document's raw OOXML.
 */
function fusedReferencesHeadingSpec(): DocSpec {
  return {
    defaultFont: "Calibri",
    defaultSizeHalf: 22,
    paragraphs: [
      { text: "The Effects of Gas Prices on Commuting", align: "center", bold: true },
      { text: "Alex Rivera", align: "center", pageBreakAfter: true },
      {
        text: "Fuel costs have risen sharply, driving changes in commuter behavior across the region.",
      },
      {
        // The bug fixture: body sentence + plain <w:br/> + bold "References" run,
        // all inside ONE <w:p>, matching the confirmed real-document shape.
        rawXml:
          `<w:p><w:r><w:t xml:space="preserve">This paper examines the shift in gas prices.</w:t></w:r>` +
          `<w:r><w:br/></w:r>` +
          `<w:r><w:rPr><w:b/></w:rPr><w:t>References</w:t></w:r></w:p>`,
      },
      {
        text: `"Fueling the mind: How local gasoline prices affect commuting choices." (2020). Journal of Transport Studies, 12(3), 55–71. https://doi.org/10.1111/jts.2020.0055`,
        indent: { firstLine: 720 },
      },
      {
        text: "Nguyen, T. (2021). Commuter response to fuel price shocks. Economics of Transportation, 8(1), 1–15. https://doi.org/10.2222/eot.2021.0001",
        indent: { firstLine: 720 },
      },
    ],
  };
}

/** A normal document whose References heading is already its own paragraph. */
function properReferencesHeadingSpec(): DocSpec {
  return {
    defaultFont: "Calibri",
    defaultSizeHalf: 22,
    paragraphs: [
      { text: "A Perfectly Ordinary Paper", align: "center", bold: true },
      { text: "Sam Lee", align: "center", pageBreakAfter: true },
      { text: "This paper discusses ordinary things at reasonable length for a body paragraph." },
      { text: "References", align: "center", bold: true, pageBreakAfter: true },
      {
        text: "Nguyen, T. (2021). Commuter response to fuel price shocks. Economics of Transportation, 8(1), 1–15. https://doi.org/10.2222/eot.2021.0001",
        indent: { firstLine: 720 },
      },
      {
        text: `"Fueling the mind." (2020). Journal of Transport Studies, 12(3), 55–71. https://doi.org/10.1111/jts.2020.0055`,
        indent: { firstLine: 720 },
      },
    ],
  };
}

/** A page-break <w:br type="page"/> fused into a paragraph — a different, already-fine shape. */
function pageBreakFusedHeadingSpec(): DocSpec {
  return {
    defaultFont: "Calibri",
    defaultSizeHalf: 22,
    paragraphs: [
      { text: "A Paper With a Page-Break Fused Heading", align: "center", bold: true },
      { text: "Sam Lee", align: "center", pageBreakAfter: true },
      { text: "This body paragraph discusses ordinary matters at reasonable length before the reference list." },
      {
        rawXml:
          `<w:p><w:r><w:t xml:space="preserve">This concludes the discussion.</w:t></w:r>` +
          `<w:r><w:br w:type="page"/></w:r>` +
          `<w:r><w:rPr><w:b/></w:rPr><w:t>References</w:t></w:r></w:p>`,
      },
      {
        text: "Nguyen, T. (2021). Commuter response to fuel price shocks. Economics of Transportation, 8(1), 1–15. https://doi.org/10.2222/eot.2021.0001",
        indent: { firstLine: 720 },
      },
    ],
  };
}

/** No references section at all. */
function noReferencesSpec(): DocSpec {
  return {
    defaultFont: "Calibri",
    defaultSizeHalf: 22,
    paragraphs: [
      { text: "A Short Paper", align: "center", bold: true },
      { text: "Sam Lee", align: "center", pageBreakAfter: true },
      { text: "This paper has no reference list at all, just body prose that goes on for a while." },
    ],
  };
}

async function runFixMode(buf: Buffer) {
  const pkg = await DocxPackage.load(buf);
  const model0 = await buildDocumentModel(pkg);
  const beforeFp = contentFingerprint(model0);
  let analysis = analyzeDocument(model0);
  let model = model0;
  let changes: import("../src/apa/types.js").Change[] = [];

  const { repairEmbeddedReferencesHeadingIfNeeded } = await import("../src/pipeline.js");
  const repaired = await repairEmbeddedReferencesHeadingIfNeeded(pkg, model, analysis);
  model = repaired.model;
  analysis = repaired.analysis;
  if (repaired.change) changes.push(repaired.change);

  const run = await runEngine(model, defaultSettings(), { fix: true, analysis });
  changes = [...changes, ...run.changes];

  const out = await pkg.save();
  await verifyDocxIntegrity(out);

  const outPkg = await DocxPackage.load(out);
  const outModel = await buildDocumentModel(outPkg);
  return { out, outModel, beforeFp, changes, model, analysis };
}

describe("embedded References-heading repair", () => {
  it("splits the fused paragraph, formats the heading, and preserves content", async () => {
    const original = await buildDocx(fusedReferencesHeadingSpec());
    const { outModel, beforeFp, changes } = await runFixMode(original);

    // The guard must not trip.
    expect(() =>
      assertContentPreserved(beforeFp, contentFingerprint(outModel), changes)
    ).not.toThrow();

    const analysis = analyzeDocument(outModel);
    expect(analysis.referencesHeadingIndex).not.toBeNull();

    const headingPara = outModel.paragraphs[analysis.referencesHeadingIndex!]!;
    expect(headingPara.text.trim()).toBe("References");
    expect(headingPara.runProps.bold).toBe(true);
    expect(headingPara.props.alignment).toBe("center");
    const newPage =
      headingPara.props.pageBreakBefore === true ||
      outModel.paragraphs[headingPara.index - 1]?.hasPageBreakAfterInRuns === true;
    expect(newPage).toBe(true);

    // Preceding body paragraph kept its sentence, minus the heading label.
    const prevPara = outModel.paragraphs[headingPara.index - 1]!;
    expect(prevPara.text.trim()).toBe("This paper examines the shift in gas prices.");

    // Reference entries get a proper 0.5" hanging indent, not a first-line indent.
    expect(analysis.referenceEntryIndexes.length).toBeGreaterThanOrEqual(2);
    for (const idx of analysis.referenceEntryIndexes) {
      const p = outModel.paragraphs[idx]!;
      expect(p.props.hangingIndent).toBe(720);
      expect(p.props.firstLineIndent ?? 0).toBe(0);
    }

    // The change log recorded the structural split.
    expect(changes.some((c) => c.ruleId === "APA-REFERENCE-000")).toBe(true);
  });

  it("check-only mode does not mutate the document but reports the issue", async () => {
    const original = await buildDocx(fusedReferencesHeadingSpec());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const beforeParagraphCount = model.paragraphs.length;
    const analysis = analyzeDocument(model);
    expect(analysis.referencesHeadingIndex).toBeNull();
    expect(analysis.embeddedReferencesHeadingCandidate).not.toBeNull();

    const audit = await runEngine(model, defaultSettings(), { fix: false, analysis });
    expect(pkg.dirtyParts.size).toBe(0);
    expect(model.paragraphs.length).toBe(beforeParagraphCount);

    const issue = audit.issues.find((i) => i.ruleId === "APA-REFERENCE-001");
    expect(issue).toBeDefined();
    expect(issue!.autoFixable).toBe(true);
    expect(issue!.message).toMatch(/line break/i);
  });

  it("leaves a document with an already-separate References paragraph untouched", async () => {
    const original = await buildDocx(properReferencesHeadingSpec());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const analysis = analyzeDocument(model);

    expect(analysis.embeddedReferencesHeadingCandidate).toBeNull();
    expect(analysis.referencesHeadingIndex).not.toBeNull();

    const { repairEmbeddedReferencesHeadingIfNeeded } = await import("../src/pipeline.js");
    const beforeCount = model.paragraphs.length;
    const repaired = await repairEmbeddedReferencesHeadingIfNeeded(pkg, model, analysis);
    expect(repaired.change).toBeNull();
    expect(repaired.model).toBe(model);
    expect(repaired.model.paragraphs.length).toBe(beforeCount);
  });

  it("does not touch a paragraph fused via a page-break (a different, already-fine pattern)", async () => {
    const original = await buildDocx(pageBreakFusedHeadingSpec());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const analysis = analyzeDocument(model);

    // The page-break case must not be picked up by the embedded-heading repair.
    expect(analysis.embeddedReferencesHeadingCandidate).toBeNull();
  });

  it("is unaffected by a document with no references section", async () => {
    const original = await buildDocx(noReferencesSpec());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const analysis = analyzeDocument(model);

    expect(analysis.referencesHeadingIndex).toBeNull();
    expect(analysis.embeddedReferencesHeadingCandidate).toBeNull();
  });
});
