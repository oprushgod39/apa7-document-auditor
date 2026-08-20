import { describe, expect, it } from "vitest";
import { DocxPackage, verifyDocxIntegrity } from "../src/docx/package.js";
import { buildDocumentModel, contentFingerprint } from "../src/docx/model.js";
import { analyzeDocument } from "../src/apa/analysis.js";
import { runEngine } from "../src/apa/engine.js";
import { assertContentPreserved } from "../src/pipeline.js";
import { defaultSettings } from "../src/apa/requirements.js";
import { buildDocx, malformedStudentPaper } from "./util/docx_builder.js";
import { TWIPS_PER_INCH } from "../src/apa/types.js";
import { childW, childrenW, descendantsW, getAttrW } from "../src/docx/xml.js";
import { findStyleEl } from "../src/docx/edit.js";

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

  it("neutralizes Word theme headings and enables contextual spacing", async () => {
    const original = await buildDocx(malformedStudentPaper());
    const { out } = await formatDocument(original);
    const pkg = await DocxPackage.load(out);
    const model = await buildDocumentModel(pkg);
    const styles = model.stylesXml!;

    const heading1 = findStyleEl(styles, "Heading1")!;
    const headingRPr = childW(heading1, "rPr")!;
    const color = childW(headingRPr, "color")!;
    expect(getAttrW(color, "val")).toBe("000000");
    expect(getAttrW(color, "themeColor")).toBeNull();
    expect(getAttrW(childW(headingRPr, "rFonts")!, "ascii")).toBe("Times New Roman");
    const headingPPr = childW(heading1, "pPr")!;
    expect(childW(headingPPr, "contextualSpacing")).not.toBeNull();
    expect(getAttrW(childW(headingPPr, "spacing")!, "after")).toBe("0");

    const normal = findStyleEl(styles, "Normal")!;
    const normalPPr = childW(normal, "pPr")!;
    expect(childW(normalPPr, "contextualSpacing")).not.toBeNull();
    expect(getAttrW(childW(normalPPr, "spacing")!, "before")).toBe("0");
    expect(getAttrW(childW(normalPPr, "spacing")!, "after")).toBe("0");

    const method = model.paragraphs.find((p) => p.text === "Method")!;
    for (const run of childrenW(method.el, "r")) {
      const runColor = childW(childW(run, "rPr"), "color")!;
      expect(getAttrW(runColor, "val")).toBe("000000");
      expect(getAttrW(runColor, "themeColor")).toBeNull();
    }
  });

  it("honors explicit heading markers and defaults unlabelled headings to Level 1", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs[7] = { text: "[H2] Participants", bold: false, align: "left" };
    spec.paragraphs.splice(
      9,
      0,
      { text: "A Clear Unlabelled Heading", bold: true, align: "left" },
      { text: "[H3] Measures", bold: false, italic: false, align: "center" }
    );
    const { out } = await formatDocument(await buildDocx(spec));
    const model = await buildDocumentModel(await DocxPackage.load(out));

    const marked = model.paragraphs.find((p) => p.text === "Participants")!;
    expect(marked).toBeDefined();
    expect(marked.styleId).toBe("Heading2");
    expect(marked.props.alignment).toBe("left");
    expect(marked.runProps.fontAscii).toBe("Times New Roman");
    expect(marked.runProps.sizeHalfPoints).toBe(24);

    const unlabelled = model.paragraphs.find((p) => p.text === "A Clear Unlabelled Heading")!;
    expect(unlabelled.styleId).toBe("Heading1");
    expect(unlabelled.props.alignment).toBe("center");
    expect(unlabelled.runProps.fontAscii).toBe("Times New Roman");
    expect(unlabelled.runProps.sizeHalfPoints).toBe(24);

    const level3 = model.paragraphs.find((p) => p.text === "Measures")!;
    expect(level3.styleId).toBe("Heading3");
    expect(level3.props.alignment).toBe("left");
    expect(level3.runProps.bold).toBe(true);
    expect(level3.runProps.italic).toBe(true);
  });

  it("removes a generic subheading prefix and starts at the first subheading level", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs[7] = { text: "Sub-heading : Level of Comparison Between AI and Gen AI" };
    const original = await buildDocx(spec);
    const before = contentFingerprint(
      await buildDocumentModel(await DocxPackage.load(original))
    );
    const { out, run } = await formatDocument(original);
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const heading = model.paragraphs.find((p) => p.text === "Level of Comparison Between AI and Gen AI")!;
    expect(heading).toBeDefined();
    expect(heading.styleId).toBe("Heading2");
    expect(heading.props.alignment).toBe("left");
    expect(heading.runProps.bold).toBe(true);
    expect(heading.runProps.fontAscii).toBe("Times New Roman");
    expect(heading.runProps.sizeHalfPoints).toBe(24);
    expect(() =>
      assertContentPreserved(before, contentFingerprint(model), run.changes)
    ).not.toThrow();
  });

  it("assigns repeated generic subheadings in order and creates Level 4/5 run-ins", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs.splice(
      7,
      0,
      { text: "Subheading: First Detail" },
      { text: "This is the body text following the first subheading and it remains a normal separate paragraph for formatting purposes." },
      { text: "Subheading: Second Detail" },
      { text: "This is the body text following the second subheading and it remains a normal separate paragraph for formatting purposes." },
      { text: "Subheading: third detail involving ai and gen ai" },
      { text: "This is the body text following the third subheading and should render on the same line after its heading." },
      { text: "Subheading: Fourth Detail" },
      { text: "This is the body text following the fourth subheading and should render on the same line after its heading." }
    );
    const original = await buildDocx(spec);
    const before = contentFingerprint(await buildDocumentModel(await DocxPackage.load(original)));
    const { out, run } = await formatDocument(original);
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const expected = [
      ["First Detail", "Heading2", false],
      ["Second Detail", "Heading3", true],
      ["Third Detail Involving AI and Gen AI.", "Heading4", false],
      ["Fourth Detail.", "Heading5", true],
    ] as const;
    for (const [text, style, italic] of expected) {
      const heading = model.paragraphs.find((p) => p.text.trimEnd() === text)!;
      expect(heading.styleId).toBe(style);
      expect(heading.runProps.bold).toBe(true);
      expect(heading.runProps.italic).toBe(italic);
    }
    for (const text of ["Third Detail Involving AI and Gen AI.", "Fourth Detail."]) {
      const heading = model.paragraphs.find((p) => p.text.trimEnd() === text)!;
      expect(
        childW(childW(childW(heading.el, "pPr"), "rPr"), "specVanish")
      ).not.toBeNull();
      const next = model.paragraphs.find((p) => p.index > heading.index && !p.isEmpty)!;
      expect(next.props.firstLineIndent ?? 0).toBe(0);
    }
    expect(() => assertContentPreserved(before, contentFingerprint(model), run.changes)).not.toThrow();
  });

  it("formats table and figure labels/titles as separate APA caption lines", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs.splice(
      9,
      0,
      { text: "Table 1", italic: true, align: "center" },
      { text: "Comparison of Results", bold: true, align: "center" },
      { text: "Figure 1", italic: true, align: "center" },
      { text: "Model Performance by Category", bold: true, align: "center" },
      { image: true }
    );
    spec.table = { rows: 2, cols: 2 };
    spec.tableAfter = 10;
    const { out } = await formatDocument(await buildDocx(spec));
    const model = await buildDocumentModel(await DocxPackage.load(out));
    for (const labelText of ["Table 1", "Figure 1"]) {
      const label = model.paragraphs.find((p) => p.text === labelText)!;
      expect(label.props.alignment).toBe("left");
      expect(label.runProps.bold).toBe(true);
      expect(label.runProps.italic ?? false).toBe(false);
    }
    for (const titleText of ["Comparison of Results", "Model Performance by Category"]) {
      const title = model.paragraphs.find((p) => p.text === titleText)!;
      expect(title.props.alignment).toBe("left");
      expect(title.runProps.bold ?? false).toBe(false);
      expect(title.runProps.italic).toBe(true);
    }
  });

  it("italicizes journal titles and volume numbers without changing reference text", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs[12]!.text = "Kumar, A., Lopez, M., & Chen, W. (2023). Does sleep architecture improve learning? Sleep Research, 12(4), 45–67. https://doi.org/10.5678/sr.2023.045";
    const original = await buildDocx(spec);
    const { out } = await formatDocument(original);
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const smith = model.paragraphs.find((p) => p.text.startsWith("Smith, J. K."))!;
    const italicText = smith.runs.filter((r) => r.effective.italic).map((r) => r.text).join("");
    expect(italicText).toContain("Journal of Cognitive Neuroscience");
    expect(italicText).toContain("36");
    expect(smith.text).toContain("Sleep deprivation and memory consolidation.");
    const kumar = model.paragraphs.find((p) => p.text.startsWith("Kumar, A."))!;
    const kumarItalic = kumar.runs.filter((r) => r.effective.italic).map((r) => r.text).join("");
    expect(kumarItalic).toContain("Sleep Research");
    expect(kumarItalic).toContain("12");
  });

  it("applies horizontal table rules, repeating headers, unsplit rows, and compact cells", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs.splice(9, 0,
      { text: "Table 1" },
      { text: "Household Effects by Group" }
    );
    spec.table = { rows: 3, cols: 3 };
    spec.tableAfter = 10;
    const { out } = await formatDocument(await buildDocx(spec));
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const table = model.tables[0]!.el;
    const borders = childW(childW(table, "tblPr"), "tblBorders")!;
    expect(getAttrW(childW(borders, "top"), "val")).toBe("single");
    expect(getAttrW(childW(borders, "bottom"), "val")).toBe("single");
    expect(getAttrW(childW(borders, "insideV"), "val")).toBe("nil");
    const rows = childrenW(table, "tr");
    expect(childW(childW(rows[0]!, "trPr"), "tblHeader")).not.toBeNull();
    for (const row of rows) {
      expect(childW(childW(row, "trPr"), "cantSplit")).not.toBeNull();
    }
    for (const p of descendantsW(table, "p")) {
      expect(getAttrW(childW(childW(p, "pPr"), "spacing"), "line")).toBe("480");
    }
  });

  it("formats figure notes flush left with only the Note label italicized", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs.splice(9, 0,
      { text: "Figure 1" },
      { text: "Inflation Burden by Income" },
      { image: true },
      { text: "Note. Values are illustrative.", indent: { firstLine: 720 } }
    );
    const { out } = await formatDocument(await buildDocx(spec));
    const model = await buildDocumentModel(await DocxPackage.load(out));
    const note = model.paragraphs.find((p) => p.text.startsWith("Note."))!;
    expect(note.props.alignment).toBe("left");
    expect(note.props.firstLineIndent ?? 0).toBe(0);
    expect(note.runs[0]!.text).toBe("Note. ");
    expect(note.runs[0]!.effective.italic).toBe(true);
    expect(note.runs.slice(1).every((r) => r.effective.italic !== true)).toBe(true);
    const audit = await runEngine(model, defaultSettings(), { fix: false, auditOnly: true });
    expect(audit.issues.find((i) => i.ruleId === "APA-LAYOUT-003")).toBeUndefined();
    expect(audit.issues.find((i) => i.ruleId === "APA-PARA-001")).toBeUndefined();
    expect(
      audit.issues.find(
        (i) => i.ruleId === "APA-HEAD-002" && /Figure 1/.test(i.message)
      )
    ).toBeUndefined();
  });
});
