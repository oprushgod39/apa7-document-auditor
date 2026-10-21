import type { Page, PDFPage, Quad, Rect } from "mupdf";

const DEFAULT_LABELS = ["Name", "University", "Subject code and name", "Professor", "Due date"];
const HEADING_EXCLUSIONS = /^(?:appendix|references|bibliography|works cited|table\s+\d+|figure\s+\d+)$/i;

export interface PdfRepairResult {
  blob: Blob;
  matches: number;
  pages: number;
  words: number;
  headingsReplaced: number;
}

interface TextLine {
  text: string;
  bbox: Rect;
  maxSize: number;
  bold: boolean;
}

interface NumberedHeading {
  rect: Rect;
  number: string;
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

export function textWithoutRanges(text: string, ranges: Array<{ start: number; end: number }>): string {
  let result = "";
  let cursor = 0;
  for (const range of ranges) {
    result += text.slice(cursor, range.start);
    if (range.start > 0 && range.end < text.length) result += " ";
    cursor = range.end;
  }
  return result + text.slice(cursor);
}

function quadBounds(quads: Quad[]): Rect {
  const xs = quads.flatMap((quad) => [quad[0], quad[2], quad[4], quad[6]]);
  const ys = quads.flatMap((quad) => [quad[1], quad[3], quad[5], quad[7]]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

function redactionRect(rect: Rect): Rect {
  // PDF glyph boxes from adjacent lines can overlap even when the printed
  // letters do not. A narrow band through the middle of the requested glyphs
  // removes those glyphs without touching the line above or below.
  const inset = Math.min(4, Math.max(0.5, (rect[3] - rect[1]) * 0.25));
  return [rect[0] - 0.1, rect[1] + inset, rect[2] + 0.1, rect[3] - inset];
}

function rectKey(rect: Rect): string {
  return rect.map((value) => Math.round(value * 10)).join(":");
}

function extractLines(page: Page): { lines: TextLine[]; text: string } {
  const structured = page.toStructuredText({});
  const lines: TextLine[] = [];
  let current: TextLine | null = null;
  structured.walk({
    beginLine(bbox) {
      current = { text: "", bbox: [...bbox] as Rect, maxSize: 0, bold: false };
    },
    onChar(character, _origin, font, size) {
      if (!current) return;
      current.text += character;
      current.maxSize = Math.max(current.maxSize, size);
      current.bold ||= font.isBold();
    },
    endLine() {
      if (current?.text.trim()) lines.push(current);
      current = null;
    }
  });
  const text = structured.asText();
  structured.destroy();
  return { lines, text };
}

function legacyHeading(lines: TextLine[], pageBounds: Rect): TextLine | null {
  const pageWidth = pageBounds[2] - pageBounds[0];
  const pageHeight = pageBounds[3] - pageBounds[1];
  const pageCenter = pageBounds[0] + pageWidth / 2;
  return lines.find((line) => {
    const text = line.text.trim();
    const center = (line.bbox[0] + line.bbox[2]) / 2;
    const nearTop = line.bbox[1] <= pageBounds[1] + pageHeight * 0.24;
    const centered = Math.abs(center - pageCenter) <= pageWidth * 0.1;
    return nearTop && centered && line.bold && line.maxSize >= 13 && !HEADING_EXCLUSIONS.test(text);
  }) ?? null;
}

function countWords(text: string): number {
  return text.match(/[\p{L}\p{N}]+(?:['’.\-][\p{L}\p{N}]+)*/gu)?.length ?? 0;
}

function addRedaction(page: PDFPage, rect: Rect): void {
  const annotation = page.createAnnotation("Redact");
  annotation.setRect(redactionRect(rect));
  annotation.setColor([1, 1, 1]);
  annotation.update();
}

function addNumber(page: PDFPage, pageBounds: Rect, heading: NumberedHeading): void {
  const annotation = page.createAnnotation("FreeText");
  const pageCenter = (pageBounds[0] + pageBounds[2]) / 2;
  annotation.setRect([
    pageCenter - 6,
    heading.rect[1] - 2,
    pageCenter + 18,
    heading.rect[3] + 4
  ]);
  annotation.setContents(heading.number);
  annotation.setDefaultAppearance("Helv", 14, [0, 0, 0]);
  annotation.setQuadding(1);
  annotation.setBorderWidth(0);
  annotation.update();
}

export async function repairPdf(
  file: File,
  terms: string[],
  removeDefaultLabels: boolean,
  replaceLegacyHeadings: boolean,
  progress: (message: string) => void
): Promise<PdfRepairResult> {
  if (file.type !== "application/pdf" && !/\.pdf$/i.test(file.name)) throw new Error("Choose a PDF file.");
  if (terms.length === 0 && !removeDefaultLabels && !replaceLegacyHeadings) {
    throw new Error("Choose at least one repair option or enter a name to remove.");
  }

  progress("Opening the original PDF text…");
  const mupdf = await import("mupdf");
  const source = mupdf.Document.openDocument(new Uint8Array(await file.arrayBuffer()), "application/pdf");
  const pdf = source.asPDF();
  if (!pdf) {
    source.destroy();
    throw new Error("This file is not a valid PDF.");
  }

  let matches = 0;
  let selectableWords = 0;
  let headingsReplaced = 0;

  try {
    for (let pageNumber = 0; pageNumber < pdf.countPages(); pageNumber++) {
      progress(`Editing page ${pageNumber + 1} of ${pdf.countPages()}…`);
      const page = pdf.loadPage(pageNumber) as PDFPage;
      const pageBounds = page.getBounds();
      const { lines, text } = extractLines(page);
      selectableWords += countWords(text);
      const redactions = new Map<string, Rect>();
      const numberedHeadings: NumberedHeading[] = [];

      for (const term of terms) {
        for (const hit of page.search(term, {})) {
          const rect = quadBounds(hit);
          redactions.set(rectKey(rect), rect);
          matches++;
        }
      }

      if (removeDefaultLabels) {
        for (const line of lines) {
          const normalized = line.text.trim().replace(/[:_\s]+$/g, "").trim();
          if (DEFAULT_LABELS.some((label) => normalized.toLocaleLowerCase() === label.toLocaleLowerCase())) {
            redactions.set(rectKey(line.bbox), line.bbox);
            matches++;
          }
        }
      }

      if (replaceLegacyHeadings) {
        const heading = legacyHeading(lines, pageBounds);
        if (heading) {
          headingsReplaced++;
          redactions.set(rectKey(heading.bbox), heading.bbox);
          numberedHeadings.push({ rect: heading.bbox, number: String(headingsReplaced) });
        }
      }

      for (const rect of redactions.values()) addRedaction(page, rect);
      if (redactions.size > 0) {
        page.applyRedactions(
          false,
          mupdf.PDFPage.REDACT_IMAGE_NONE,
          mupdf.PDFPage.REDACT_LINE_ART_NONE,
          mupdf.PDFPage.REDACT_TEXT_REMOVE
        );
      }
      for (const heading of numberedHeadings) addNumber(page, pageBounds, heading);
      page.destroy();
      await new Promise((resolve) => globalThis.setTimeout(resolve, 0));
    }

    if (selectableWords === 0) {
      throw new Error("This PDF contains no real text. It is an image-only scan and needs OCR before words can be edited.");
    }
    if (matches === 0 && headingsReplaced === 0) {
      throw new Error("No matching names, labels, or old document headings were found.");
    }

    if (headingsReplaced > 0) pdf.bake(true, false);
    progress("Verifying searchable words…");
    const saved = pdf.saveToBuffer("garbage,compress").asUint8Array();
    const verification = mupdf.Document.openDocument(saved, "application/pdf");
    let searchableText = "";
    try {
      for (let index = 0; index < verification.countPages(); index++) {
        const page = verification.loadPage(index);
        const structured = page.toStructuredText({});
        searchableText += `\n${structured.asText()}`;
        structured.destroy();
        page.destroy();
      }
    } finally {
      verification.destroy();
    }
    const words = countWords(searchableText);
    if (words === 0) throw new Error("The PDF repair did not preserve searchable words. No output was created.");
    const remainingTerms = terms.filter((term) => findRedactionRanges(searchableText, [term]).length > 0);
    if (remainingTerms.length > 0) {
      throw new Error("One or more requested names could still be extracted, so no output was created. Check the spelling and try again.");
    }

    const copied = Uint8Array.from(saved);
    return {
      blob: new Blob([copied.buffer], { type: "application/pdf" }),
      matches,
      pages: pdf.countPages(),
      words,
      headingsReplaced
    };
  } finally {
    source.destroy();
  }
}
