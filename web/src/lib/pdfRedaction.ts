import { PDFDocument, StandardFonts, type PDFFont, type PDFPage } from "pdf-lib";

const DEFAULT_LABELS = ["Name", "University", "Subject code and name", "Professor", "Due date"];

export interface PdfRepairResult {
  blob: Blob;
  matches: number;
  pages: number;
}

export function parseRedactionTerms(value: string): string[] {
  return [...new Set(value.split(/[\n,;]+/).map((term) => term.trim()).filter((term) => term.length >= 2))];
}

export function findRedactionRanges(text: string, terms: string[]): Array<{ start: number; end: number }> {
  const lower = text.toLocaleLowerCase();
  const ranges: Array<{ start: number; end: number }> = [];
  const isWord = (character: string | undefined) => character != null && /[\p{L}\p{N}]/u.test(character);
  for (const term of terms) {
    const needle = term.toLocaleLowerCase();
    let from = 0;
    while (from < lower.length) {
      const start = lower.indexOf(needle, from);
      if (start < 0) break;
      const end = start + needle.length;
      if (!isWord(lower[start - 1]) && !isWord(lower[end])) ranges.push({ start, end });
      from = Math.max(end, start + 1);
    }
  }
  ranges.sort((left, right) => left.start - right.start || left.end - right.end);
  return ranges.filter((range, index) => index === 0 || range.start >= ranges[index - 1]!.end);
}

function defaultLabelRanges(text: string): Array<{ start: number; end: number }> {
  const trimmed = text.trim();
  const normalized = trimmed.replace(/[:_\s]+$/g, "").trim().toLocaleLowerCase();
  return DEFAULT_LABELS.some((label) => normalized === label.toLocaleLowerCase())
    ? [{ start: text.indexOf(trimmed), end: text.indexOf(trimmed) + trimmed.length }]
    : [];
}

function mergeRanges(ranges: Array<{ start: number; end: number }>): Array<{ start: number; end: number }> {
  const sorted = ranges.slice().sort((left, right) => left.start - right.start || left.end - right.end);
  const merged: Array<{ start: number; end: number }> = [];
  for (const range of sorted) {
    const previous = merged[merged.length - 1];
    if (!previous || range.start > previous.end) merged.push({ ...range });
    else previous.end = Math.max(previous.end, range.end);
  }
  return merged;
}

export function textWithoutRanges(text: string, ranges: Array<{ start: number; end: number }>): string {
  let result = "";
  let cursor = 0;
  for (const range of mergeRanges(ranges)) {
    result += text.slice(cursor, range.start);
    // Keep a separator so words on either side of removed content do not join together.
    if (range.start > 0 && range.end < text.length) result += " ";
    cursor = range.end;
  }
  return result + text.slice(cursor);
}

function searchableWinAnsiText(text: string, font: PDFFont): string {
  let result = "";
  for (const character of text) {
    try {
      font.encodeText(character);
      result += character;
    } catch {
      result += " ";
    }
  }
  return result;
}

function addSearchableText(
  page: PDFPage,
  font: PDFFont,
  text: string,
  transform: number[]
): void {
  const searchable = searchableWinAnsiText(text, font);
  if (!searchable.trim()) return;
  const size = Math.max(1, Math.hypot(transform[2] ?? 0, transform[3] ?? 0));
  page.drawText(searchable, {
    x: transform[4] ?? 0,
    y: transform[5] ?? 0,
    size,
    font,
    opacity: 0
  });
}

function canvasBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("The repaired page could not be encoded.")), "image/jpeg", 0.94);
  });
}

export async function repairPdf(
  file: File,
  terms: string[],
  removeDefaultLabels: boolean,
  progress: (message: string) => void
): Promise<PdfRepairResult> {
  if (file.type !== "application/pdf" && !/\.pdf$/i.test(file.name)) throw new Error("Choose a PDF file.");
  if (terms.length === 0 && !removeDefaultLabels) throw new Error("Enter at least one student or university name to remove.");

  const pdfjsLib = await import("pdfjs-dist");
  const pdfWorker = (await import("pdfjs-dist/build/pdf.worker.min.mjs?url")).default;
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;
  const source = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const output = await PDFDocument.create();
  const searchableFont = await output.embedFont(StandardFonts.Helvetica);
  let matches = 0;
  let selectableItems = 0;

  for (let pageNumber = 1; pageNumber <= source.numPages; pageNumber++) {
    progress(`Repairing page ${pageNumber} of ${source.numPages}…`);
    const page = await source.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2 });
    const original = page.getViewport({ scale: 1 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("This browser cannot prepare the repaired PDF.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvas, canvasContext: context, viewport }).promise;

    const content = await page.getTextContent();
    const searchableItems: Array<{ text: string; transform: number[] }> = [];
    for (const raw of content.items) {
      if (!("str" in raw) || !raw.str.trim() || !("transform" in raw) || !("width" in raw)) continue;
      selectableItems++;
      const ranges = mergeRanges([
        ...findRedactionRanges(raw.str, terms),
        ...(removeDefaultLabels ? defaultLabelRanges(raw.str) : [])
      ]);
      searchableItems.push({ text: textWithoutRanges(raw.str, ranges), transform: Array.from(raw.transform) });
      if (ranges.length === 0) continue;
      const transformed = pdfjsLib.Util.transform(viewport.transform, raw.transform);
      const fontHeight = Math.max(8, Math.hypot(transformed[2], transformed[3]));
      const itemWidth = Math.max(fontHeight, raw.width * 2);
      for (const range of ranges) {
        const startRatio = range.start / Math.max(1, raw.str.length);
        const endRatio = range.end / Math.max(1, raw.str.length);
        const x = transformed[4] + itemWidth * startRatio - 3;
        const width = itemWidth * (endRatio - startRatio) + 6;
        const y = transformed[5] - fontHeight * 1.08;
        context.fillStyle = "#ffffff";
        context.fillRect(x, y, width, fontHeight * 1.35);
        matches++;
      }
    }

    const image = await output.embedJpg(await (await canvasBlob(canvas)).arrayBuffer());
    const outputPage = output.addPage([original.width, original.height]);
    outputPage.drawImage(image, { x: 0, y: 0, width: original.width, height: original.height });
    // The rendered page preserves appearance. This invisible text layer restores
    // search, copy/paste, accessibility, and word counting without reintroducing
    // any of the redacted names or labels.
    for (const item of searchableItems) addSearchableText(outputPage, searchableFont, item.text, item.transform);
    canvas.width = 1;
    canvas.height = 1;
    await new Promise((resolve) => window.setTimeout(resolve, 0));
  }

  if (selectableItems === 0) throw new Error("This PDF has no selectable text. Scanned image-only PDFs need OCR before automatic redaction.");
  if (matches === 0) throw new Error("No matching names or title-page labels were found. Enter the text exactly as it appears in the PDF.");

  progress("Preparing download…");
  const bytes = await output.save({ useObjectStreams: true });
  const copied = Uint8Array.from(bytes);
  return { blob: new Blob([copied.buffer], { type: "application/pdf" }), matches, pages: source.numPages };
}
