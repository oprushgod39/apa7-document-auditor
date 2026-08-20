import { DocxPackage } from "../docx/package.js";
import { buildDocumentModel, type ParagraphModel, type TableModel } from "../docx/model.js";
import { AppError } from "../errors.js";
import { MergePdfRenderer } from "./pdf_render.js";
import { buildAppendixParagraphs, REFERENCE_HEADING } from "./word_count.js";

export interface MergeInput {
  name: string;
  originalName: string;
  buffer: Buffer;
}

type Block =
  | { kind: "paragraph"; blockIndex: number; paragraph: ParagraphModel }
  | { kind: "table"; blockIndex: number; table: TableModel };

/**
 * Merges 2-30 source .docx files (in the given order, each preceded by a
 * centered user-supplied heading) into a single PDF. Each source document
 * has its References/Bibliography/Works Cited section (and everything
 * after it) stripped before being included — the cutoff is located by
 * paragraph blockIndex using the exact same REFERENCE_HEADING detection
 * word_count.ts uses, so this matches the word counts /merge-preview
 * already promised the user. A controlled appendix is appended last.
 *
 * The whole pipeline is pure JavaScript: the app's existing OOXML model
 * (DocxPackage/buildDocumentModel, already proven in production for the
 * main APA-auditor feature) is parsed directly and drawn into the PDF with
 * pdfkit. There is no headless browser, no native binary, and no
 * serverless-vs-local branching — the exact same code path runs in dev, in
 * tests, and on Vercel.
 */
export async function mergeDocuments(inputs: MergeInput[], appendixWords: number): Promise<Buffer> {
  const renderer = new MergePdfRenderer();

  try {
    for (const input of inputs) {
      let pkg: DocxPackage;
      let model: Awaited<ReturnType<typeof buildDocumentModel>>;
      try {
        pkg = await DocxPackage.load(input.buffer);
        model = await buildDocumentModel(pkg);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        throw new AppError(
          "PROCESSING_FAILED",
          `"${input.originalName}" could not be read for merging (${message}).`,
          422
        );
      }

      const reference = model.paragraphs.find((p) => REFERENCE_HEADING.test(p.text.trim()));
      const cutoff = reference?.blockIndex ?? Number.POSITIVE_INFINITY;

      const blocks: Block[] = [
        ...model.paragraphs
          .filter((p) => p.blockIndex < cutoff)
          .map((paragraph): Block => ({ kind: "paragraph", blockIndex: paragraph.blockIndex, paragraph })),
        ...model.tables
          .filter((t) => t.blockIndex < cutoff)
          .map((table): Block => ({ kind: "table", blockIndex: table.blockIndex, table })),
      ].sort((a, b) => a.blockIndex - b.blockIndex);

      renderer.startSection();
      renderer.drawHeading(input.name);
      for (const block of blocks) {
        if (block.kind === "paragraph") {
          await renderer.drawParagraph(pkg, block.paragraph);
        } else {
          renderer.drawTable(block.table);
        }
      }
    }

    const appendixParagraphs = await buildAppendixParagraphs(appendixWords);
    if (appendixParagraphs.length > 0) {
      renderer.startSection();
      renderer.drawHeading("Appendix");
      for (const paragraph of appendixParagraphs) {
        renderer.drawAppendixParagraph(paragraph.text, paragraph.bold);
      }
    }

    return await renderer.finish();
  } catch (error) {
    if (error instanceof AppError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new AppError("PROCESSING_FAILED", `The documents could not be merged (${message}).`, 500);
  }
}
