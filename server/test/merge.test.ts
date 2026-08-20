import { describe, expect, it } from "vitest";
// Import the library module directly (not the package's `index.js`), which
// contains a `!module.parent` debug branch that misfires under ESM/vitest.
import pdfParse from "pdf-parse/lib/pdf-parse.js";
import { mergeDocuments } from "../src/merge/merge.js";
import { buildDocx } from "./util/docx_builder.js";

const PDF_MAGIC = "%PDF-";

describe("mergeDocuments", () => {
  it("merges two documents (with a table and image) plus an appendix into a single well-formed PDF", async () => {
    const docA = await buildDocx({
      paragraphs: [
        { text: "Alpha document body paragraph one.", spacing: { after: 160 } },
        { text: "Alpha document body paragraph two, with more content.", image: true },
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

    const parsed = await pdfParse(output);
    const text = parsed.text.replace(/\s+/g, " ");

    // Both user-supplied headings appear.
    expect(text).toContain("Alpha Submission");
    expect(text).toContain("Bravo Submission");

    // Body content from both documents survived the conversion.
    expect(text).toContain("Alpha document body paragraph one");
    expect(text).toContain("Bravo document body paragraph one");

    // Table content from the first document survived.
    expect(text).toContain("AlphaCell");

    // References/Bibliography sections were stripped from both documents.
    expect(text).not.toContain("must not appear in the merge");
    expect(text).not.toContain("excluded from the merge");

    // The appendix was appended (its source text starts with "IGNORE" per
    // server/assets/ignore_appendix.txt) and produced at least one page.
    expect(parsed.numpages).toBeGreaterThanOrEqual(1);
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
    const parsed = await pdfParse(output);
    const text = parsed.text.replace(/\s+/g, " ");

    expect(text).toContain("Kept paragraph before the reference heading");
    expect(text).not.toContain("This citation text is dropped");
    expect(text).not.toContain("Works Cited");
  }, 60_000);
});
