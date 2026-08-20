import { AppError } from "../errors.js";
import { docxToTrimmedHtml } from "./html_convert.js";
import { renderHtmlToPdf } from "./pdf_render.js";
import { buildAppendixParagraphs, type AppendixParagraph } from "./word_count.js";

export interface MergeInput {
  name: string;
  originalName: string;
  buffer: Buffer;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderSection(heading: string, bodyHtml: string, isFirst: boolean): string {
  const pageBreak = isFirst ? "" : ' style="page-break-before: always;"';
  return `<section class="merge-doc"${pageBreak}><p class="merge-heading">${escapeHtml(heading)}</p>${bodyHtml}</section>`;
}

function renderAppendix(paragraphs: AppendixParagraph[]): string {
  if (paragraphs.length === 0) return "";
  const body = paragraphs
    .map((p) => (p.bold ? `<p><strong>${escapeHtml(p.text)}</strong></p>` : `<p>${escapeHtml(p.text)}</p>`))
    .join("");
  return `<section class="merge-doc" style="page-break-before: always;"><p class="merge-heading">Appendix</p>${body}</section>`;
}

function wrapDocument(sectionsHtml: string): string {
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page { margin: 1in; }
  body {
    font-family: "Times New Roman", Georgia, serif;
    font-size: 12pt;
    line-height: 1.6;
    color: #000;
  }
  p { margin: 0 0 8pt 0; }
  table { border-collapse: collapse; margin: 8pt 0; }
  td, th { border: 1px solid #444; padding: 4pt 6pt; }
  img { max-width: 100%; }
  .merge-heading {
    text-align: center;
    font-weight: bold;
    font-size: 14pt;
    margin-bottom: 12pt;
  }
  .merge-doc:first-child { margin-top: 0; }
</style>
</head>
<body>
${sectionsHtml}
</body>
</html>`;
}

/**
 * Merges 2-30 source .docx files (in the given order, each preceded by a
 * centered user-supplied heading) into a single PDF. Each source document
 * has its References/Bibliography/Works Cited section (and everything after
 * it) stripped before being included, mirroring countDocumentWords' cutoff
 * detection exactly (see html_convert.ts). Formatting, tables, and images
 * are preserved as faithfully as HTML/PDF rendering allows; a controlled
 * appendix is appended last.
 *
 * Runs entirely in Node — no Microsoft Word, PowerShell, or Windows
 * dependency — so it works both locally and on Vercel's Linux serverless
 * runtime.
 */
export async function mergeDocuments(inputs: MergeInput[], appendixWords: number): Promise<Buffer> {
  const sectionsHtml: string[] = [];
  for (let index = 0; index < inputs.length; index++) {
    const input = inputs[index]!;
    let trimmedHtml: string;
    try {
      trimmedHtml = await docxToTrimmedHtml(input.buffer);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new AppError(
        "PROCESSING_FAILED",
        `"${input.originalName}" could not be converted for merging (${message}).`,
        422
      );
    }
    sectionsHtml.push(renderSection(input.name, trimmedHtml, index === 0));
  }

  const appendixParagraphs = await buildAppendixParagraphs(appendixWords);
  sectionsHtml.push(renderAppendix(appendixParagraphs));

  const fullHtml = wrapDocument(sectionsHtml.join(""));

  try {
    return await renderHtmlToPdf(fullHtml);
  } catch (error) {
    if (error instanceof AppError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new AppError("PROCESSING_FAILED", `The documents could not be merged (${message}).`, 500);
  }
}
