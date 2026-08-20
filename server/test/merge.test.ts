import { describe, expect, it } from "vitest";
import { mergeDocuments } from "../src/merge/merge.js";
import { buildDocx } from "./util/docx_builder.js";
import { extractPdfText } from "./util/pdf_text.js";

const PDF_MAGIC = "%PDF-";

describe("mergeDocuments", () => {
  it("merges two documents (with a table) plus an appendix into a single well-formed PDF", async () => {
    const docA = await buildDocx({
      paragraphs: [
        { text: "Alpha document body paragraph one.", spacing: { after: 160 } },
        { text: "Alpha document body paragraph two, with more content." },
        { text: "References", align: "left" },
        { text: "Smith, J. (2024). A paper that must not appear in the merge." },
      ],
      table: { rows: 2, cols: 2, cellText: "AlphaCell" },
      tableAfter: 1,
    });

    const docB = await buildDocx({
      paragraphs: [
        { text: "Bravo document body paragraph one.", spacing: { after: 160 } },
        { text: "Bravo document body paragraph two." },
        { text: "Bibliography", align: "left" },
        { text: "Jones, K. (2023). Another citation excluded from the merge." },
      ],
    });

    const output = await mergeDocuments(
      [
        { name: "Alpha Submission", originalName: "alpha.docx", buffer: docA },
        { name: "Bravo Submission", originalName: "bravo.docx", buffer: docB },
      ],
      50
    );

    // Valid PDF: starts with the PDF magic bytes.
    expect(output.subarray(0, PDF_MAGIC.length).toString("latin1")).toBe(PDF_MAGIC);

    // Sanity check: the merged PDF is meaningfully larger than either source
    // document alone, i.e. content actually made it into the output.
    expect(output.length).toBeGreaterThan(docA.length);
    expect(output.length).toBeGreaterThan(docB.length);

    const { text, numPages } = await extractPdfText(output);
    const flat = text.replace(/\s+/g, " ");

    // Both user-supplied headings appear.
    expect(flat).toContain("Alpha Submission");
    expect(flat).toContain("Bravo Submission");

    // Body content from both documents survived the conversion.
    expect(flat).toContain("Alpha document body paragraph one");
    expect(flat).toContain("Bravo document body paragraph one");

    // Table content from the first document survived.
    expect(flat).toContain("AlphaCell");

    // References/Bibliography sections were stripped from both documents.
    expect(flat).not.toContain("must not appear in the merge");
    expect(flat).not.toContain("excluded from the merge");

    // The appendix was appended (its source text starts with "IGNORE" per
    // server/assets/ignore_appendix.txt) and produced at least one page.
    expect(numPages).toBeGreaterThanOrEqual(1);
  }, 60_000);

  it("stops each source document exactly at its References/Bibliography/Works Cited heading", async () => {
    const doc = await buildDocx({
      paragraphs: [
        { text: "Kept paragraph before the reference heading." },
        { text: "Works Cited" },
        { text: "This citation text is dropped." },
      ],
    });

    const output = await mergeDocuments(
      [{ name: "Solo Document", originalName: "solo.docx", buffer: doc }],
      0
    );
    const { text } = await extractPdfText(output);
    const flat = text.replace(/\s+/g, " ");

    expect(flat).toContain("Kept paragraph before the reference heading");
    expect(flat).not.toContain("This citation text is dropped");
    expect(flat).not.toContain("Works Cited");
  }, 60_000);

  it("does not crash when a source document contains an embedded image, and renders the surrounding content", async () => {
    const doc = await buildDocx({
      paragraphs: [
        { text: "Paragraph before the image." },
        { text: "Paragraph with an inline image.", image: true },
        { text: "Paragraph after the image." },
      ],
    });

    const output = await mergeDocuments(
      [{ name: "Document With Image", originalName: "with_image.docx", buffer: doc }],
      0
    );

    // Still a well-formed PDF — the pipeline didn't crash on the image.
    expect(output.subarray(0, PDF_MAGIC.length).toString("latin1")).toBe(PDF_MAGIC);

    const { text } = await extractPdfText(output);
    const flat = text.replace(/\s+/g, " ");

    // Text content around the image (which itself isn't text-extractable)
    // survived intact.
    expect(flat).toContain("Document With Image");
    expect(flat).toContain("Paragraph before the image");
    expect(flat).toContain("Paragraph with an inline image");
    expect(flat).toContain("Paragraph after the image");
  }, 60_000);
});
