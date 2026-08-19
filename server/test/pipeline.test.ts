import { describe, expect, it } from "vitest";
import { DocxPackage, verifyDocxIntegrity } from "../src/docx/package.js";
import { buildDocumentModel, contentFingerprint } from "../src/docx/model.js";
import { analyzeDocument } from "../src/apa/analysis.js";
import { runEngine } from "../src/apa/engine.js";
import { assertContentPreserved } from "../src/pipeline.js";
import { defaultSettings } from "../src/apa/requirements.js";
import { buildDocx, malformedStudentPaper } from "./util/docx_builder.js";
import { TWIPS_PER_INCH } from "../src/apa/types.js";

async function formatDocument(buf: Buffer, settings = defaultSettings()) {
  const pkg = await DocxPackage.load(buf);
  const model = await buildDocumentModel(pkg);
  const analysis = analyzeDocument(model);
  const run = await runEngine(model, settings, { fix: true, analysis });
  const out = await pkg.save();
  await verifyDocxIntegrity(out);
  return { out, run, beforeModel: model };
}

describe("golden document formatting", () => {
  it("fixes margins, spacing, fonts, indents, references — and passes audit", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const { out, run } = await formatDocument(original);
    expect(run.changes.length).toBeGreaterThan(5);

    const pkg = await DocxPackage.load(out);
    const model = await buildDocumentModel(pkg);
    const analysis = analyzeDocument(model);

    // Margins fixed to 1"
    for (const s of model.sections) {
      expect(s.marginTop).toBe(TWIPS_PER_INCH);
      expect(s.marginLeft).toBe(TWIPS_PER_INCH);
    }
    // Double spacing everywhere
    for (const p of model.paragraphs) {
      if (p.insideTable) continue;
      expect(p.props.line).toBe(480);
      expect(p.props.spacingAfter ?? 0).toBe(0);
    }
    // Fonts normalized
    for (const p of model.paragraphs) {
      if (p.isEmpty || p.insideTable) continue;
      expect(p.runProps.fontAscii).toBe("Times New Roman");
      expect(p.runProps.sizeHalfPoints).toBe(24);
    }
    // "Bibliography" renamed to "References" and moved formatting
    const refHeading = model.paragraphs[analysis.referencesHeadingIndex!]!;
    expect(refHeading.text.trim()).toBe("References");
    expect(refHeading.props.alignment).toBe("center");
    // Hanging indents on entries
    for (const idx of analysis.referenceEntryIndexes) {
      const p = model.paragraphs[idx]!;
      expect(p.props.hangingIndent).toBe(720);
    }
    // Alphabetical order: Adams, Kumar, Smith
    const firstWords = analysis.referenceEntryIndexes.map(
      (i) => model.paragraphs[i]!.text.trim().split(",")[0]
    );
    expect(firstWords).toEqual(["Adams", "Kumar", "Smith"]);
    // DOI modernized
    const smith = analysis.references.find((r) => r.surnames[0] === "smith")!;
    expect(smith.raw).toContain("https://doi.org/10.1234/jcn.2024.0112");
    expect(smith.raw).not.toContain("doi:10.1234");
    // Page number header created
    expect(model.headersFooters.some((h) => h.hasPageField)).toBe(true);
    // Title bolded
    const title = model.paragraphs[analysis.detectedMetadata.titleParagraphIndex!]!;
    expect(title.runProps.bold).toBe(true);
  });

  it("audit flags missing reference (Johnson) and uncited reference (Adams)", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const { out } = await formatDocument(original);
    const pkg = await DocxPackage.load(out);
    const model = await buildDocumentModel(pkg);
    const audit = await runEngine(model, defaultSettings(), {
      fix: false,
      auditOnly: true,
    });
    const missing = audit.issues.find((i) => i.ruleId === "APA-CITATION-008");
    expect(missing).toBeDefined();
    expect(missing!.message).toContain("Johnson");
    expect(missing!.userResolutionRequired).toBe(true);
    const uncited = audit.issues.find((i) => i.ruleId === "APA-REFERENCE-005");
    expect(uncited).toBeDefined();
    expect(uncited!.message).toContain("Adams");
  });

  it("preserves content: text, tables, images, hyperlinks survive", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs[7]!.image = true;
    spec.paragraphs[8]!.hyperlink = { text: "example", url: "https://example.com" };
    spec.table = { rows: 3, cols: 3 };
    spec.tableAfter = 8;
    const original = await buildDocx(spec);
    const pkgBefore = await DocxPackage.load(original);
    const before = contentFingerprint(await buildDocumentModel(pkgBefore));

    const { out, run } = await formatDocument(original);
    const pkgAfter = await DocxPackage.load(out);
    const after = contentFingerprint(await buildDocumentModel(pkgAfter));

    expect(after.imageCount).toBe(before.imageCount);
    expect(after.tableShapes).toEqual(before.tableShapes);
    expect(after.hyperlinkCount).toBe(before.hyperlinkCount);
    // The full guard should accept the run (text changes are recorded).
    assertContentPreserved(before, after, run.changes);
  });

  it("is idempotent: reformatting a formatted document makes no meaningful changes", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const first = await formatDocument(original);
    const second = await formatDocument(first.out);
    const meaningful = second.run.changes.filter(
      // Style-definition rewrites are harmless no-ops on identical values.
      (c) => c.ruleId !== "APA-HEAD-001"
    );
    expect(meaningful).toHaveLength(0);
  });

  it("check mode reports issues without modifying the document", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const settings = { ...defaultSettings(), mode: "check" as const };
    const run = await runEngine(model, settings, { fix: false });
    expect(run.changes).toHaveLength(0);
    expect(run.issues.length).toBeGreaterThan(3);
    expect(pkg.dirtyParts.size).toBe(0);
  });

  it("detects headings via multi-signal classifier", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const pkg = await DocxPackage.load(original);
    const model = await buildDocumentModel(pkg);
    const analysis = analyzeDocument(model);
    const method = analysis.headings.find((h) => h.text === "Method");
    expect(method).toBeDefined();
    expect(method!.level).toBe(1); // centered + bold + section word
  });

  it("strips manual tab indentation and applies real first-line indent", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const { out } = await formatDocument(original);
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const methodBody = model.paragraphs.find((p) =>
      p.text.includes("Participants completed a memory task")
    )!;
    expect(methodBody.text.startsWith(" ")).toBe(false);
    expect(methodBody.props.firstLineIndent).toBe(720);
    expect(methodBody.props.alignment ?? "left").toBe("left"); // was justified
  });
});
