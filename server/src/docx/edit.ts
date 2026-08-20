import type { DocxPackage } from "./package.js";
import {
  NS,
  childW,
  childrenW,
  createW,
  ensureChildW,
  getAttrW,
  removeAttrW,
  removeChildW,
  setAttrW,
  type XDocument,
  type XElement,
} from "./xml.js";

/**
 * OOXML property containers are order-sensitive (CT_PPr / CT_RPr sequences).
 * These canonical orders let us insert new children at schema-valid positions
 * without disturbing existing content.
 */
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
  "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
  "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
  "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
  "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
  "suppressOverlap", "jc", "textDirection", "textAlignment",
  "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
];

const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
  "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
  "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
  "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
  "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
  "specVanish", "oMath",
];

function afterNames(order: string[], local: string): string[] {
  const i = order.indexOf(local);
  return i >= 0 ? order.slice(i + 1) : [];
}

function ensureOrdered(
  doc: XDocument,
  parent: XElement,
  local: string,
  order: string[]
): XElement {
  return ensureChildW(doc, parent, local, afterNames(order, local));
}

export function ensurePPr(doc: XDocument, p: XElement): XElement {
  const existing = childW(p, "pPr");
  if (existing) return existing;
  const pPr = createW(doc, "pPr");
  p.insertBefore(pPr, p.firstChild);
  return pPr;
}

function ensurePPrChild(doc: XDocument, p: XElement, local: string): XElement {
  return ensureOrdered(doc, ensurePPr(doc, p), local, PPR_ORDER);
}

export function ensureRPr(doc: XDocument, r: XElement): XElement {
  const existing = childW(r, "rPr");
  if (existing) return existing;
  const rPr = createW(doc, "rPr");
  r.insertBefore(rPr, r.firstChild);
  return rPr;
}

// --- Paragraph-level edits -------------------------------------------------

export interface SpacingSpec {
  before?: number;
  after?: number;
  line?: number;
  lineRule?: string;
}

export function setParagraphSpacing(
  doc: XDocument,
  p: XElement,
  spec: SpacingSpec
): void {
  const spacing = ensurePPrChild(doc, p, "spacing");
  if (spec.before !== undefined) setAttrW(spacing, "before", String(spec.before));
  if (spec.after !== undefined) setAttrW(spacing, "after", String(spec.after));
  if (spec.line !== undefined) setAttrW(spacing, "line", String(spec.line));
  if (spec.lineRule !== undefined) setAttrW(spacing, "lineRule", spec.lineRule);
  // "auto" spacing flags override numeric values; clear them.
  removeAttrW(spacing, "beforeAutospacing");
  removeAttrW(spacing, "afterAutospacing");
}

/** Enable Word's “Don't add space between paragraphs of the same style”. */
export function setParagraphContextualSpacing(
  doc: XDocument,
  p: XElement,
  on = true
): void {
  const pPr = ensurePPr(doc, p);
  if (on) {
    const contextual = ensureOrdered(doc, pPr, "contextualSpacing", PPR_ORDER);
    removeAttrW(contextual, "val");
  } else {
    removeChildW(pPr, "contextualSpacing");
  }
}

export interface IndentSpec {
  firstLine?: number | null; // null = remove attribute
  hanging?: number | null;
  left?: number | null;
}

export function setParagraphIndent(
  doc: XDocument,
  p: XElement,
  spec: IndentSpec
): void {
  const ind = ensurePPrChild(doc, p, "ind");
  const apply = (attr: string, v: number | null | undefined) => {
    if (v === undefined) return;
    if (v === null) removeAttrW(ind, attr);
    else setAttrW(ind, attr, String(v));
  };
  // firstLine and hanging are mutually exclusive in OOXML.
  if (spec.firstLine !== undefined && spec.firstLine !== null) {
    removeAttrW(ind, "hanging");
  }
  if (spec.hanging !== undefined && spec.hanging !== null) {
    removeAttrW(ind, "firstLine");
  }
  apply("firstLine", spec.firstLine);
  apply("hanging", spec.hanging);
  apply("left", spec.left);
  if (spec.left !== undefined) removeAttrW(ind, "start");
  if (!ind.attributes || ind.attributes.length === 0) {
    ind.parentNode?.removeChild(ind);
  }
}

export function setParagraphAlignment(
  doc: XDocument,
  p: XElement,
  value: string
): void {
  const jc = ensurePPrChild(doc, p, "jc");
  setAttrW(jc, "val", value);
}

export function setParagraphStyle(
  doc: XDocument,
  p: XElement,
  styleId: string
): void {
  const pStyle = ensurePPrChild(doc, p, "pStyle");
  setAttrW(pStyle, "val", styleId);
}

export function removeParagraphStyle(p: XElement): void {
  const pPr = childW(p, "pPr");
  if (pPr) removeChildW(pPr, "pStyle");
}

export function setPageBreakBefore(doc: XDocument, p: XElement, on: boolean): void {
  const pPr = ensurePPr(doc, p);
  if (on) {
    ensureOrdered(doc, pPr, "pageBreakBefore", PPR_ORDER);
  } else {
    removeChildW(pPr, "pageBreakBefore");
  }
}

export function setParagraphKeepNext(doc: XDocument, p: XElement, on: boolean): void {
  const pPr = ensurePPr(doc, p);
  if (on) {
    const keepNext = ensureOrdered(doc, pPr, "keepNext", PPR_ORDER);
    removeAttrW(keepNext, "val");
  } else {
    removeChildW(pPr, "keepNext");
  }
}

export function setParagraphKeepLines(doc: XDocument, p: XElement, on: boolean): void {
  const pPr = ensurePPr(doc, p);
  if (on) {
    const keepLines = ensureOrdered(doc, pPr, "keepLines", PPR_ORDER);
    removeAttrW(keepLines, "val");
  } else {
    removeChildW(pPr, "keepLines");
  }
}

/**
 * Turn the paragraph mark into a Word style separator. The following
 * paragraph remains structurally separate but continues on the same rendered
 * line, which is required for APA Levels 4 and 5.
 */
export function setParagraphStyleSeparator(
  doc: XDocument,
  p: XElement,
  on: boolean
): void {
  const pPr = ensurePPr(doc, p);
  const markRPr = ensureOrdered(doc, pPr, "rPr", PPR_ORDER);
  if (on) {
    // Word writes both flags for a style-separator paragraph mark. `vanish`
    // hides the mark in layout while `specVanish` identifies why it is hidden.
    const vanish = ensureOrdered(doc, markRPr, "vanish", RPR_ORDER);
    removeAttrW(vanish, "val");
    const specVanish = ensureOrdered(doc, markRPr, "specVanish", RPR_ORDER);
    removeAttrW(specVanish, "val");
  } else {
    removeChildW(markRPr, "vanish");
    removeChildW(markRPr, "specVanish");
    if (!markRPr.firstChild) pPr.removeChild(markRPr);
  }
}

/** Replace a plain-text paragraph with explicitly formatted runs. */
export function replaceParagraphRuns(
  doc: XDocument,
  p: XElement,
  segments: { text: string; bold?: boolean; italic?: boolean; font?: string; halfPoints?: number; black?: boolean }[]
): boolean {
  if (childrenW(p, "hyperlink").length > 0) return false;
  if (p.getElementsByTagNameNS(NS.w, "drawing").length > 0) return false;
  for (const r of childrenW(p, "r")) p.removeChild(r);
  for (const segment of segments) {
    if (!segment.text) continue;
    const r = createW(doc, "r");
    if (segment.font) setRunFont(doc, r, segment.font);
    if (segment.halfPoints) setRunSize(doc, r, segment.halfPoints);
    if (segment.bold !== undefined) setRunBold(doc, r, segment.bold);
    if (segment.italic !== undefined) setRunItalic(doc, r, segment.italic);
    if (segment.black) setRunColorBlack(doc, r);
    const t = createW(doc, "t");
    t.setAttribute("xml:space", "preserve");
    t.appendChild(doc.createTextNode(segment.text));
    r.appendChild(t);
    p.appendChild(r);
  }
  return true;
}

/** Apply APA's minimal horizontal-rule table layout and stable pagination. */
export function formatApaTable(
  doc: XDocument,
  table: XElement,
  opts: { font: string; halfPoints: number }
): void {
  const tblPr = childW(table, "tblPr") ?? (() => {
    const el = createW(doc, "tblPr");
    table.insertBefore(el, table.firstChild);
    return el;
  })();
  removeChildW(tblPr, "tblBorders");
  const borders = ensureChildW(doc, tblPr, "tblBorders", [
    "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
  ]);
  const border = (name: string, val: "single" | "nil") => {
    const el = createW(doc, name);
    setAttrW(el, "val", val);
    if (val === "single") {
      setAttrW(el, "sz", "8");
      setAttrW(el, "space", "0");
      setAttrW(el, "color", "000000");
    }
    borders.appendChild(el);
  };
  border("top", "single");
  border("left", "nil");
  border("bottom", "single");
  border("right", "nil");
  border("insideH", "nil");
  border("insideV", "nil");

  const rows = childrenW(table, "tr");
  rows.forEach((row, rowIndex) => {
    const trPr = childW(row, "trPr") ?? (() => {
      const el = createW(doc, "trPr");
      row.insertBefore(el, row.firstChild);
      return el;
    })();
    const cantSplit = ensureChildW(doc, trPr, "cantSplit");
    removeAttrW(cantSplit, "val");
    if (rowIndex === 0) {
      const header = ensureChildW(doc, trPr, "tblHeader");
      removeAttrW(header, "val");
    }
    childrenW(row, "tc").forEach((cell, colIndex) => {
      if (rowIndex === 0) {
        const tcPr = childW(cell, "tcPr") ?? (() => {
          const el = createW(doc, "tcPr");
          cell.insertBefore(el, cell.firstChild);
          return el;
        })();
        const tcBorders = childW(tcPr, "tcBorders") ??
          ensureChildW(doc, tcPr, "tcBorders", ["shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign"]);
        removeChildW(tcBorders, "bottom");
        const bottom = createW(doc, "bottom");
        setAttrW(bottom, "val", "single");
        setAttrW(bottom, "sz", "8");
        setAttrW(bottom, "space", "0");
        setAttrW(bottom, "color", "000000");
        tcBorders.appendChild(bottom);
      }
      for (const p of childrenW(cell, "p")) {
        setParagraphSpacing(doc, p, { before: 0, after: 0, line: 480, lineRule: "auto" });
        setParagraphContextualSpacing(doc, p);
        setParagraphIndent(doc, p, { firstLine: null, hanging: null, left: 0 });
        setParagraphAlignment(doc, p, rowIndex === 0 && colIndex > 0 ? "center" : "left");
        setParagraphKeepLines(doc, p, true);
        setParagraphKeepNext(doc, p, rowIndex === 0);
        setParagraphRunFonts(doc, p, opts.font, opts.halfPoints);
      }
    });
  });
}

// --- Run-level edits -------------------------------------------------------

export function setRunFont(doc: XDocument, r: XElement, font: string): void {
  const rPr = ensureRPr(doc, r);
  const rFonts = ensureOrdered(doc, rPr, "rFonts", RPR_ORDER);
  setAttrW(rFonts, "ascii", font);
  setAttrW(rFonts, "hAnsi", font);
  // leave eastAsia / cs alone unless present and latin-like
  const cs = getAttrW(rFonts, "cs");
  if (cs) setAttrW(rFonts, "cs", font);
}

export function setRunSize(doc: XDocument, r: XElement, halfPoints: number): void {
  const rPr = ensureRPr(doc, r);
  const sz = ensureOrdered(doc, rPr, "sz", RPR_ORDER);
  setAttrW(sz, "val", String(halfPoints));
  const szCs = ensureOrdered(doc, rPr, "szCs", RPR_ORDER);
  setAttrW(szCs, "val", String(halfPoints));
}

export function setRunBold(doc: XDocument, r: XElement, on: boolean): void {
  const rPr = ensureRPr(doc, r);
  if (on) {
    ensureOrdered(doc, rPr, "b", RPR_ORDER).removeAttributeNS(NS.w, "val");
    ensureOrdered(doc, rPr, "bCs", RPR_ORDER).removeAttributeNS(NS.w, "val");
  } else {
    setAttrW(ensureOrdered(doc, rPr, "b", RPR_ORDER), "val", "0");
    setAttrW(ensureOrdered(doc, rPr, "bCs", RPR_ORDER), "val", "0");
  }
}

export function setRunItalic(doc: XDocument, r: XElement, on: boolean): void {
  const rPr = ensureRPr(doc, r);
  if (on) {
    ensureOrdered(doc, rPr, "i", RPR_ORDER).removeAttributeNS(NS.w, "val");
    ensureOrdered(doc, rPr, "iCs", RPR_ORDER).removeAttributeNS(NS.w, "val");
  } else {
    setAttrW(ensureOrdered(doc, rPr, "i", RPR_ORDER), "val", "0");
    setAttrW(ensureOrdered(doc, rPr, "iCs", RPR_ORDER), "val", "0");
  }
}

/** Force a run to ordinary black text and remove all theme-color inheritance. */
export function setRunColorBlack(doc: XDocument, r: XElement): void {
  const rPr = ensureRPr(doc, r);
  const color = ensureOrdered(doc, rPr, "color", RPR_ORDER);
  setAttrW(color, "val", "000000");
  removeAttrW(color, "themeColor");
  removeAttrW(color, "themeTint");
  removeAttrW(color, "themeShade");
}

export function setRunUnderlineNone(doc: XDocument, r: XElement): void {
  const rPr = ensureRPr(doc, r);
  const underline = ensureOrdered(doc, rPr, "u", RPR_ORDER);
  setAttrW(underline, "val", "none");
  removeAttrW(underline, "color");
  removeAttrW(underline, "themeColor");
}

export function setParagraphRunColorBlack(doc: XDocument, p: XElement): void {
  for (const r of childrenW(p, "r")) setRunColorBlack(doc, r);
  for (const h of childrenW(p, "hyperlink")) {
    for (const r of childrenW(h, "r")) setRunColorBlack(doc, r);
  }
}

/** Apply font/size to every run in a paragraph plus its paragraph mark. */
export function setParagraphRunFonts(
  doc: XDocument,
  p: XElement,
  font: string,
  halfPoints: number
): void {
  for (const r of childrenW(p, "r")) {
    setRunFont(doc, r, font);
    setRunSize(doc, r, halfPoints);
  }
  // Also fix runs inside hyperlinks.
  for (const h of childrenW(p, "hyperlink")) {
    for (const r of childrenW(h, "r")) {
      setRunFont(doc, r, font);
      setRunSize(doc, r, halfPoints);
    }
  }
  const pPr = childW(p, "pPr");
  const markRPr = childW(pPr, "rPr");
  if (markRPr) {
    const rFonts = childW(markRPr, "rFonts");
    if (rFonts) {
      setAttrW(rFonts, "ascii", font);
      setAttrW(rFonts, "hAnsi", font);
    }
    const sz = childW(markRPr, "sz");
    if (sz) setAttrW(sz, "val", String(halfPoints));
  }
}

// --- Style edits (styles.xml) ---------------------------------------------

export function findStyleEl(stylesDoc: XDocument, styleId: string): XElement | null {
  const root = stylesDoc.documentElement as XElement;
  for (const s of childrenW(root, "style")) {
    if (getAttrW(s, "styleId") === styleId) return s;
  }
  return null;
}

export function setStyleRunFormatting(
  stylesDoc: XDocument,
  styleEl: XElement,
  spec: {
    font?: string;
    halfPoints?: number;
    bold?: boolean;
    italic?: boolean;
    black?: boolean;
  }
): void {
  const rPr = ensureChildW(stylesDoc, styleEl, "rPr", []);
  if (spec.font !== undefined) {
    const rFonts = ensureOrdered(stylesDoc, rPr, "rFonts", RPR_ORDER);
    setAttrW(rFonts, "ascii", spec.font);
    setAttrW(rFonts, "hAnsi", spec.font);
  }
  if (spec.halfPoints !== undefined) {
    setAttrW(ensureOrdered(stylesDoc, rPr, "sz", RPR_ORDER), "val", String(spec.halfPoints));
    setAttrW(ensureOrdered(stylesDoc, rPr, "szCs", RPR_ORDER), "val", String(spec.halfPoints));
  }
  if (spec.bold !== undefined) {
    if (spec.bold) {
      ensureOrdered(stylesDoc, rPr, "b", RPR_ORDER).removeAttributeNS(NS.w, "val");
    } else {
      removeChildW(rPr, "b");
    }
  }
  if (spec.italic !== undefined) {
    if (spec.italic) {
      ensureOrdered(stylesDoc, rPr, "i", RPR_ORDER).removeAttributeNS(NS.w, "val");
    } else {
      removeChildW(rPr, "i");
    }
  }
  if (spec.black) {
    const color = ensureOrdered(stylesDoc, rPr, "color", RPR_ORDER);
    setAttrW(color, "val", "000000");
    removeAttrW(color, "themeColor");
    removeAttrW(color, "themeTint");
    removeAttrW(color, "themeShade");
  }
}

export function setStyleParaFormatting(
  stylesDoc: XDocument,
  styleEl: XElement,
  spec: SpacingSpec & {
    alignment?: string;
    firstLine?: number | null;
    contextualSpacing?: boolean;
    keepNext?: boolean;
  }
): void {
  const rPrEl = childW(styleEl, "rPr");
  const pPr = (() => {
    const existing = childW(styleEl, "pPr");
    if (existing) return existing;
    const el = createW(stylesDoc, "pPr");
    // pPr must come before rPr in a style definition.
    if (rPrEl) styleEl.insertBefore(el, rPrEl);
    else styleEl.appendChild(el);
    return el;
  })();
  if (
    spec.before !== undefined ||
    spec.after !== undefined ||
    spec.line !== undefined ||
    spec.lineRule !== undefined
  ) {
    const spacing = ensureOrdered(stylesDoc, pPr, "spacing", PPR_ORDER);
    if (spec.before !== undefined) setAttrW(spacing, "before", String(spec.before));
    if (spec.after !== undefined) setAttrW(spacing, "after", String(spec.after));
    if (spec.line !== undefined) setAttrW(spacing, "line", String(spec.line));
    if (spec.lineRule !== undefined) setAttrW(spacing, "lineRule", spec.lineRule);
  }
  if (spec.alignment !== undefined) {
    setAttrW(ensureOrdered(stylesDoc, pPr, "jc", PPR_ORDER), "val", spec.alignment);
  }
  if (spec.firstLine !== undefined) {
    const ind = ensureOrdered(stylesDoc, pPr, "ind", PPR_ORDER);
    if (spec.firstLine === null) {
      removeAttrW(ind, "firstLine");
    } else {
      setAttrW(ind, "firstLine", String(spec.firstLine));
      removeAttrW(ind, "hanging");
    }
  }
  if (spec.contextualSpacing !== undefined) {
    if (spec.contextualSpacing) {
      const contextual = ensureOrdered(stylesDoc, pPr, "contextualSpacing", PPR_ORDER);
      removeAttrW(contextual, "val");
    } else {
      removeChildW(pPr, "contextualSpacing");
    }
  }
  if (spec.keepNext !== undefined) {
    if (spec.keepNext) {
      const keepNext = ensureOrdered(stylesDoc, pPr, "keepNext", PPR_ORDER);
      removeAttrW(keepNext, "val");
    } else {
      removeChildW(pPr, "keepNext");
    }
  }
}

/** Create a paragraph style if it does not exist. Returns the style element. */
export function ensureParagraphStyle(
  stylesDoc: XDocument,
  styleId: string,
  name: string,
  basedOn = "Normal"
): XElement {
  const existing = findStyleEl(stylesDoc, styleId);
  if (existing) return existing;
  const root = stylesDoc.documentElement as XElement;
  const style = createW(stylesDoc, "style");
  setAttrW(style, "type", "paragraph");
  setAttrW(style, "styleId", styleId);
  const nameEl = createW(stylesDoc, "name");
  setAttrW(nameEl, "val", name);
  style.appendChild(nameEl);
  if (findStyleEl(stylesDoc, basedOn)) {
    const basedOnEl = createW(stylesDoc, "basedOn");
    setAttrW(basedOnEl, "val", basedOn);
    style.appendChild(basedOnEl);
  }
  const qFormat = createW(stylesDoc, "qFormat");
  style.appendChild(qFormat);
  root.appendChild(style);
  return style;
}

// --- Structural edits ------------------------------------------------------

/** Create a new paragraph with a single run of text. */
export function createParagraph(
  doc: XDocument,
  text: string,
  opts: {
    styleId?: string;
    alignment?: string;
    bold?: boolean;
    italic?: boolean;
    black?: boolean;
    font?: string;
    halfPoints?: number;
    spacing?: SpacingSpec;
    firstLineIndent?: number | null;
  } = {}
): XElement {
  const p = createW(doc, "p");
  if (opts.styleId) setParagraphStyle(doc, p, opts.styleId);
  if (opts.alignment) setParagraphAlignment(doc, p, opts.alignment);
  if (opts.spacing) setParagraphSpacing(doc, p, opts.spacing);
  if (opts.firstLineIndent !== undefined) {
    setParagraphIndent(doc, p, { firstLine: opts.firstLineIndent });
  }
  if (text.length > 0) {
    const r = createW(doc, "r");
    if (opts.bold || opts.italic || opts.font || opts.halfPoints || opts.black) {
      const rPr = createW(doc, "rPr");
      r.appendChild(rPr);
      if (opts.font) {
        const rFonts = createW(doc, "rFonts");
        setAttrW(rFonts, "ascii", opts.font);
        setAttrW(rFonts, "hAnsi", opts.font);
        rPr.appendChild(rFonts);
      }
      if (opts.bold) rPr.appendChild(createW(doc, "b"));
      if (opts.italic) rPr.appendChild(createW(doc, "i"));
      if (opts.black) {
        const color = createW(doc, "color");
        setAttrW(color, "val", "000000");
        rPr.appendChild(color);
      }
      if (opts.halfPoints) {
        const sz = createW(doc, "sz");
        setAttrW(sz, "val", String(opts.halfPoints));
        rPr.appendChild(sz);
      }
    }
    const t = createW(doc, "t");
    t.setAttribute("xml:space", "preserve");
    t.appendChild(doc.createTextNode(text));
    r.appendChild(t);
    p.appendChild(r);
  }
  return p;
}

export function insertBeforeEl(parent: XElement, newEl: XElement, ref: XElement | null): void {
  if (ref) parent.insertBefore(newEl, ref);
  else parent.appendChild(newEl);
}

// --- Section / page setup --------------------------------------------------

export function setSectionMargins(
  doc: XDocument,
  sectPr: XElement,
  twips: { top: number; bottom: number; left: number; right: number }
): void {
  // sectPr child order: headerReference*, footerReference*, footnotePr,
  // endnotePr, type, pgSz, pgMar, ...
  let pgMar = childW(sectPr, "pgMar");
  if (!pgMar) {
    pgMar = createW(doc, "pgMar");
    const pgSz = childW(sectPr, "pgSz");
    if (pgSz && pgSz.nextSibling) sectPr.insertBefore(pgMar, pgSz.nextSibling);
    else sectPr.appendChild(pgMar);
  }
  setAttrW(pgMar, "top", String(twips.top));
  setAttrW(pgMar, "bottom", String(twips.bottom));
  setAttrW(pgMar, "left", String(twips.left));
  setAttrW(pgMar, "right", String(twips.right));
  if (!getAttrW(pgMar, "header")) setAttrW(pgMar, "header", "720");
  if (!getAttrW(pgMar, "footer")) setAttrW(pgMar, "footer", "720");
}

// --- Page-number header creation ------------------------------------------

const HEADER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const HEADER_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

function headerXml(font: string, halfPoints: number, leftText?: string): string {
  const esc = (s: string) =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const rPr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr>`;
  const runningHead = leftText
    ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(leftText)}</w:t></w:r><w:r>${rPr}<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>`
    : "";
  const align = leftText ? "" : `<w:jc w:val="right"/>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<w:hdr xmlns:w="${NS.w}" xmlns:r="${NS.r}">` +
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${align}<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${halfPoints}"/></w:rPr></w:pPr>` +
    runningHead +
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>` +
    `</w:p></w:hdr>`
  );
}

/**
 * Ensure every section has a default header containing a right-aligned PAGE
 * field (and optionally a running head on the left). Creates the header part,
 * content-type override and relationship if needed.
 */
export async function ensurePageNumberHeader(
  pkg: DocxPackage,
  documentXml: XDocument,
  sectPrs: XElement[],
  opts: { font: string; halfPoints: number; runningHead?: string }
): Promise<{ created: boolean; partName: string | null }> {
  const sectionsNeeding = sectPrs.filter(
    (s) => !childrenW(s, "headerReference").some((r) => getAttrW(r, "type") === "default")
  );
  if (sectionsNeeding.length === 0) return { created: false, partName: null };

  // Pick a non-colliding header part name.
  let n = 1;
  while (pkg.has(`word/header${n}.xml`)) n++;
  const partName = `word/header${n}.xml`;
  const partFile = `header${n}.xml`;

  pkg.setTextPart(
    partName,
    headerXml(opts.font, opts.halfPoints, opts.runningHead)
  );

  // Content type override
  const ctDoc = await pkg.getXml("[Content_Types].xml");
  const ctRoot = ctDoc.documentElement as XElement;
  const override = ctDoc.createElementNS(NS.ct, "Override");
  override.setAttribute("PartName", `/${partName}`);
  override.setAttribute("ContentType", HEADER_CONTENT_TYPE);
  ctRoot.appendChild(override);
  pkg.markDirty("[Content_Types].xml");

  // Relationship
  let relsDoc = await pkg.getXmlIfPresent("word/_rels/document.xml.rels");
  if (!relsDoc) {
    pkg.setTextPart(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rels}"/>`
    );
    relsDoc = await pkg.getXml("word/_rels/document.xml.rels");
  }
  const relsRoot = relsDoc.documentElement as XElement;
  const existing = relsRoot.getElementsByTagNameNS(NS.rels, "Relationship");
  let maxId = 0;
  for (let i = 0; i < existing.length; i++) {
    const id = (existing.item(i) as XElement).getAttribute("Id") ?? "";
    const m = /^rId(\d+)$/.exec(id);
    if (m) maxId = Math.max(maxId, Number.parseInt(m[1]!, 10));
  }
  const rId = `rId${maxId + 1}`;
  const rel = relsDoc.createElementNS(NS.rels, "Relationship");
  rel.setAttribute("Id", rId);
  rel.setAttribute("Type", HEADER_REL_TYPE);
  rel.setAttribute("Target", partFile);
  relsRoot.appendChild(rel);
  pkg.markDirty("word/_rels/document.xml.rels");

  // headerReference in each section lacking one — must be first children of sectPr.
  for (const sectPr of sectionsNeeding) {
    const ref = documentXml.createElementNS(NS.w, "w:headerReference");
    setAttrW(ref, "type", "default");
    ref.setAttributeNS(NS.r, "r:id", rId);
    sectPr.insertBefore(ref, sectPr.firstChild);
  }
  pkg.markDirty("word/document.xml");
  return { created: true, partName };
}
