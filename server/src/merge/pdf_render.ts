import PDFDocument from "pdfkit";
import type { DocxPackage } from "../docx/package.js";
import type { ParagraphModel, TableModel } from "../docx/model.js";
import { childrenW, paragraphText } from "../docx/xml.js";
import { extractRunImages } from "./images.js";

const PAGE_MARGIN = 72; // 1 inch, matches the previous CSS @page margin
const LETTER_WIDTH = 612; // 8.5in at 72pt/in — pdfkit's "LETTER" page size
const BODY_FONT_SIZE = 12;
const TABLE_FONT_SIZE = 10.5;
const TABLE_CELL_PADDING = 4;
const TABLE_MIN_ROW_HEIGHT = 18;

function fontFor(bold?: boolean, italic?: boolean): string {
  if (bold && italic) return "Times-BoldItalic";
  if (bold) return "Times-Bold";
  if (italic) return "Times-Italic";
  return "Times-Roman";
}

function mapAlign(alignment?: string): "left" | "center" | "right" | "justify" {
  switch (alignment) {
    case "center":
      return "center";
    case "right":
      return "right";
    case "both":
      return "justify";
    default:
      return "left";
  }
}

/**
 * Draws the merged submission PDF directly with pdfkit — pure JavaScript,
 * no native binary, no headless browser. Every page-break and layout
 * decision (paragraph flow, table row pagination, image fit) is computed
 * explicitly, since pdfkit has no built-in HTML/CSS layout engine.
 */
export class MergePdfRenderer {
  private readonly doc: PDFKit.PDFDocument;
  private readonly chunks: Buffer[] = [];
  private readonly done: Promise<Buffer>;
  private readonly contentWidth: number;

  constructor() {
    this.doc = new PDFDocument({
      size: "LETTER",
      margins: { top: PAGE_MARGIN, bottom: PAGE_MARGIN, left: PAGE_MARGIN, right: PAGE_MARGIN },
      autoFirstPage: false,
      bufferPages: true,
    });
    // this.doc.page doesn't exist until the first addPage() call
    // (autoFirstPage is false), so compute from the fixed page size instead.
    this.contentWidth = LETTER_WIDTH - PAGE_MARGIN * 2;
    this.doc.on("data", (chunk: Buffer) => this.chunks.push(chunk));
    this.done = new Promise<Buffer>((resolve, reject) => {
      this.doc.on("end", () => resolve(Buffer.concat(this.chunks)));
      this.doc.on("error", reject);
    });
  }

  /** Starts a new page for the next source document (or the appendix). */
  startSection(): void {
    this.doc.addPage();
    this.doc.font("Times-Roman").fontSize(BODY_FONT_SIZE);
  }

  drawHeading(text: string): void {
    this.doc.x = this.doc.page.margins.left;
    this.doc.font("Times-Bold").fontSize(14).text(text, { align: "center" });
    this.doc.moveDown(1);
    this.doc.font("Times-Roman").fontSize(BODY_FONT_SIZE);
    this.doc.x = this.doc.page.margins.left;
  }

  async drawParagraph(pkg: DocxPackage, paragraph: ParagraphModel): Promise<void> {
    const doc = this.doc;
    doc.x = doc.page.margins.left;

    if (paragraph.isEmpty && !paragraph.hasDrawing) {
      doc.font("Times-Roman").fontSize(BODY_FONT_SIZE).moveDown(0.6);
      return;
    }

    const runsWithText = paragraph.runs.filter((r) => r.text.length > 0);
    const renderRuns =
      runsWithText.length > 0
        ? runsWithText.map((r) => ({ text: r.text, bold: r.effective.bold, italic: r.effective.italic }))
        : paragraph.text.trim().length > 0
          ? [{ text: paragraph.text, bold: paragraph.runProps.bold, italic: paragraph.runProps.italic }]
          : [];

    if (renderRuns.length > 0) {
      const align = mapAlign(paragraph.props.alignment);
      doc.fontSize(BODY_FONT_SIZE);
      renderRuns.forEach((run, idx) => {
        doc.font(fontFor(run.bold, run.italic));
        doc.text(run.text, { continued: idx < renderRuns.length - 1, align });
      });
    }

    if (paragraph.hasDrawing) {
      for (const run of paragraph.runs) {
        if (!run.hasDrawing) continue;
        const images = await extractRunImages(pkg, run.el);
        for (const image of images) this.drawImage(image);
      }
    }

    doc.font("Times-Roman").fontSize(BODY_FONT_SIZE).moveDown(0.6);
    doc.x = doc.page.margins.left;
  }

  drawAppendixParagraph(text: string, bold: boolean): void {
    const doc = this.doc;
    doc.x = doc.page.margins.left;
    if (text.trim().length === 0) {
      doc.font("Times-Roman").fontSize(BODY_FONT_SIZE).moveDown(0.6);
      return;
    }
    doc
      .font(fontFor(bold, false))
      .fontSize(BODY_FONT_SIZE)
      .text(text, { align: "left" });
    doc.font("Times-Roman").fontSize(BODY_FONT_SIZE).moveDown(0.6);
    doc.x = doc.page.margins.left;
  }

  /**
   * Draws an image scaled to the content width (preserving aspect ratio).
   * If it won't reasonably fit in the remaining space on the current page,
   * a new page is started first rather than letting it clip or overlap
   * following content.
   */
  private drawImage(buffer: Buffer): void {
    const doc = this.doc;
    let image: { width: number; height: number };
    try {
      // openImage() is a pdfkit runtime API not covered by @types/pdfkit.
      image = (doc as unknown as { openImage(src: Buffer): { width: number; height: number } }).openImage(
        buffer
      );
    } catch {
      return; // unsupported/corrupt image data — skip rather than crash
    }
    if (!image.width || !image.height) return;

    const aspect = image.width / image.height;
    const top = doc.page.margins.top;
    const bottomLimit = doc.page.height - doc.page.margins.bottom;
    const fullPageHeight = bottomLimit - top;

    let width = this.contentWidth;
    let height = width / aspect;

    let remaining = bottomLimit - doc.y;
    if (height > remaining && doc.y > top) {
      doc.addPage();
      doc.font("Times-Roman").fontSize(BODY_FONT_SIZE);
      remaining = fullPageHeight;
    }
    if (height > remaining) {
      height = remaining;
      width = height * aspect;
      if (width > this.contentWidth) {
        width = this.contentWidth;
        height = width / aspect;
      }
    }

    const x = doc.page.margins.left + (this.contentWidth - width) / 2;
    const y = doc.y;
    doc.image(buffer, x, y, { width, height });
    doc.y = y + height + 8;
    doc.x = doc.page.margins.left;
  }

  /**
   * Minimal table layout: equal-width columns spanning the content width,
   * text-wrapped within cells, thin vector borders, automatic row height
   * from wrapped text height, and page-break-aware row rendering (a row
   * that would run past the bottom margin starts a fresh page before it is
   * drawn, so it never straddles a page break).
   */
  drawTable(table: TableModel): void {
    const doc = this.doc;
    const rows = childrenW(table.el, "tr");
    if (rows.length === 0) return;
    const cols = Math.max(1, table.cols);
    const left = doc.page.margins.left;
    const colWidth = this.contentWidth / cols;

    doc.font("Times-Roman").fontSize(TABLE_FONT_SIZE);
    doc.moveDown(0.3);

    for (const tr of rows) {
      const cells = childrenW(tr, "tc");
      const texts: string[] = [];
      for (let c = 0; c < cols; c++) {
        const tc = cells[c];
        let text = "";
        if (tc) {
          for (const p of childrenW(tc, "p")) {
            const t = paragraphText(p).trim();
            if (t) text += (text ? "\n" : "") + t;
          }
        }
        texts.push(text);
      }

      let rowHeight = TABLE_MIN_ROW_HEIGHT;
      for (const text of texts) {
        const h =
          doc.heightOfString(text || " ", { width: colWidth - TABLE_CELL_PADDING * 2 }) +
          TABLE_CELL_PADDING * 2;
        if (h > rowHeight) rowHeight = h;
      }

      const bottomLimit = doc.page.height - doc.page.margins.bottom;
      if (doc.y + rowHeight > bottomLimit) {
        doc.addPage();
        doc.font("Times-Roman").fontSize(TABLE_FONT_SIZE);
      }

      const top = doc.y;
      doc.lineWidth(0.5).strokeColor("#444444");
      for (let c = 0; c < cols; c++) {
        const x = left + c * colWidth;
        doc.rect(x, top, colWidth, rowHeight).stroke();
        doc
          .fillColor("#000000")
          .text(texts[c] ?? "", x + TABLE_CELL_PADDING, top + TABLE_CELL_PADDING, {
            width: colWidth - TABLE_CELL_PADDING * 2,
          });
      }
      doc.x = left;
      doc.y = top + rowHeight;
    }

    doc.font("Times-Roman").fontSize(BODY_FONT_SIZE).moveDown(0.5);
    doc.x = left;
  }

  async finish(): Promise<Buffer> {
    this.doc.end();
    return this.done;
  }
}
