import type { DocxPackage } from "./package.js";
import {
  NS,
  childW,
  childrenW,
  descendantsW,
  getAttrW,
  paragraphText,
  type XDocument,
  type XElement,
} from "./xml.js";

/** Effective run (character) properties after style resolution. */
export interface RunProps {
  fontAscii?: string;
  fontHAnsi?: string;
  sizeHalfPoints?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  allCaps?: boolean;
}

/** Effective paragraph properties after style resolution. Twips units. */
export interface ParaProps {
  alignment?: string; // left | center | right | both | ...
  spacingBefore?: number;
  spacingAfter?: number;
  line?: number;
  lineRule?: string; // auto | exact | atLeast
  firstLineIndent?: number;
  hangingIndent?: number;
  leftIndent?: number;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
}

export interface StyleInfo {
  id: string;
  name: string;
  type: string;
  basedOn?: string;
  isDefault: boolean;
  rPr: RunProps;
  pPr: ParaProps;
}

export interface RunModel {
  el: XElement;
  text: string;
  props: RunProps; // direct formatting only
  effective: RunProps; // after style resolution
  hasPageBreak: boolean;
  hasDrawing: boolean;
}

export interface ParagraphModel {
  /** Index within all body paragraphs (tables excluded), 0-based. */
  index: number;
  /** Index within body block sequence (paragraphs + tables). */
  blockIndex: number;
  el: XElement;
  styleId: string | null;
  styleName: string | null;
  text: string;
  isEmpty: boolean;
  /** Direct paragraph formatting only. */
  direct: ParaProps;
  /** Effective formatting: docDefaults <- style chain <- direct. */
  props: ParaProps;
  /** Effective run props of the dominant (longest) run, or paragraph mark. */
  runProps: RunProps;
  runs: RunModel[];
  hasDrawing: boolean;
  hasPageBreakAfterInRuns: boolean;
  /** True if paragraph belongs to a numbered/bulleted list. */
  hasNumbering: boolean;
  /** True if this paragraph contains a sectPr (end-of-section marker). */
  hasSectPr: boolean;
  insideTable: boolean;
}

export interface SectionModel {
  el: XElement;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  pageWidth?: number;
  pageHeight?: number;
  headerRefs: { type: string; rId: string }[];
  footerRefs: { type: string; rId: string }[];
  titlePg: boolean;
  pgNumStart?: number;
}

export interface TableModel {
  index: number;
  blockIndex: number;
  el: XElement;
  rows: number;
  cols: number;
  /** Text of first cell — useful for caption/number checks. */
  firstCellText: string;
}

export interface HeaderFooterModel {
  partName: string;
  kind: "header" | "footer";
  refType: string; // default | first | even
  text: string;
  hasPageField: boolean;
}

export interface DocumentModel {
  pkg: DocxPackage;
  documentXml: XDocument;
  stylesXml: XDocument | null;
  body: XElement;
  paragraphs: ParagraphModel[];
  tables: TableModel[];
  sections: SectionModel[];
  styles: Map<string, StyleInfo>;
  styleIdByName: Map<string, string>;
  docDefaults: { rPr: RunProps; pPr: ParaProps };
  headersFooters: HeaderFooterModel[];
  footnoteCount: number;
  endnoteCount: number;
  imageCount: number;
  hyperlinkCount: number;
  equationCount: number;
}

// ---------------------------------------------------------------------------

function parseRunProps(rPr: XElement | null): RunProps {
  const out: RunProps = {};
  if (!rPr) return out;
  const rFonts = childW(rPr, "rFonts");
  if (rFonts) {
    const ascii = getAttrW(rFonts, "ascii");
    const hAnsi = getAttrW(rFonts, "hAnsi");
    if (ascii) out.fontAscii = ascii;
    if (hAnsi) out.fontHAnsi = hAnsi;
  }
  const sz = childW(rPr, "sz");
  if (sz) {
    const v = getAttrW(sz, "val");
    if (v) out.sizeHalfPoints = Number.parseInt(v, 10);
  }
  const flag = (name: string): boolean | undefined => {
    const el = childW(rPr, name);
    if (!el) return undefined;
    const v = getAttrW(el, "val");
    return v == null ? true : v !== "0" && v !== "false" && v !== "none";
  };
  const b = flag("b");
  if (b !== undefined) out.bold = b;
  const i = flag("i");
  if (i !== undefined) out.italic = i;
  const u = childW(rPr, "u");
  if (u) out.underline = getAttrW(u, "val") !== "none";
  const caps = flag("caps");
  if (caps !== undefined) out.allCaps = caps;
  return out;
}

function parseParaProps(pPr: XElement | null): ParaProps {
  const out: ParaProps = {};
  if (!pPr) return out;
  const jc = childW(pPr, "jc");
  if (jc) {
    const v = getAttrW(jc, "val");
    if (v) out.alignment = v;
  }
  const spacing = childW(pPr, "spacing");
  if (spacing) {
    const before = getAttrW(spacing, "before");
    const after = getAttrW(spacing, "after");
    const line = getAttrW(spacing, "line");
    const lineRule = getAttrW(spacing, "lineRule");
    if (before != null) out.spacingBefore = Number.parseInt(before, 10);
    if (after != null) out.spacingAfter = Number.parseInt(after, 10);
    if (line != null) out.line = Number.parseInt(line, 10);
    if (lineRule != null) out.lineRule = lineRule;
  }
  const ind = childW(pPr, "ind");
  if (ind) {
    const firstLine = getAttrW(ind, "firstLine");
    const hanging = getAttrW(ind, "hanging");
    const left = getAttrW(ind, "left") ?? getAttrW(ind, "start");
    if (firstLine != null) out.firstLineIndent = Number.parseInt(firstLine, 10);
    if (hanging != null) out.hangingIndent = Number.parseInt(hanging, 10);
    if (left != null) out.leftIndent = Number.parseInt(left, 10);
  }
  if (childW(pPr, "contextualSpacing")) out.contextualSpacing = true;
  if (childW(pPr, "pageBreakBefore")) out.pageBreakBefore = true;
  return out;
}

function mergeRun(base: RunProps, over: RunProps): RunProps {
  return { ...base, ...over };
}

function mergePara(base: ParaProps, over: ParaProps): ParaProps {
  return { ...base, ...over };
}

function parseStyles(stylesXml: XDocument | null): {
  styles: Map<string, StyleInfo>;
  byName: Map<string, string>;
  docDefaults: { rPr: RunProps; pPr: ParaProps };
} {
  const styles = new Map<string, StyleInfo>();
  const byName = new Map<string, string>();
  let defaults = { rPr: {} as RunProps, pPr: {} as ParaProps };
  if (!stylesXml || !stylesXml.documentElement) {
    return { styles, byName, docDefaults: defaults };
  }
  const root = stylesXml.documentElement as XElement;
  const dd = childW(root, "docDefaults");
  if (dd) {
    const rDef = childW(childW(dd, "rPrDefault"), "rPr");
    const pDef = childW(childW(dd, "pPrDefault"), "pPr");
    defaults = { rPr: parseRunProps(rDef), pPr: parseParaProps(pDef) };
  }
  for (const styleEl of childrenW(root, "style")) {
    const id = getAttrW(styleEl, "styleId");
    if (!id) continue;
    const type = getAttrW(styleEl, "type") ?? "paragraph";
    const nameEl = childW(styleEl, "name");
    const name = (nameEl && getAttrW(nameEl, "val")) || id;
    const basedOnEl = childW(styleEl, "basedOn");
    const basedOn = basedOnEl ? getAttrW(basedOnEl, "val") ?? undefined : undefined;
    const isDefault = getAttrW(styleEl, "default") === "1";
    styles.set(id, {
      id,
      name,
      type,
      basedOn,
      isDefault,
      rPr: parseRunProps(childW(styleEl, "rPr")),
      pPr: parseParaProps(childW(styleEl, "pPr")),
    });
    if (type === "paragraph") byName.set(name.toLowerCase(), id);
  }
  return { styles, byName, docDefaults: defaults };
}

function resolveStyleChain(
  styleId: string | null,
  styles: Map<string, StyleInfo>
): StyleInfo[] {
  const chain: StyleInfo[] = [];
  const seen = new Set<string>();
  let cur = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const s = styles.get(cur);
    if (!s) break;
    chain.unshift(s); // base first
    cur = s.basedOn ?? null;
  }
  return chain;
}

function fieldInstructions(root: XElement): string[] {
  const out: string[] = [];
  for (const instr of descendantsW(root, "instrText")) {
    out.push(instr.textContent ?? "");
  }
  for (const fld of descendantsW(root, "fldSimple")) {
    const v = getAttrW(fld, "instr");
    if (v) out.push(v);
  }
  return out;
}

// ---------------------------------------------------------------------------

export async function buildDocumentModel(pkg: DocxPackage): Promise<DocumentModel> {
  const documentXml = await pkg.getXml("word/document.xml");
  const stylesXml = await pkg.getXmlIfPresent("word/styles.xml");
  const { styles, byName, docDefaults } = parseStyles(stylesXml);

  const docRoot = documentXml.documentElement as XElement;
  const body = childW(docRoot, "body");
  if (!body) throw new Error("document.xml has no w:body");

  const paragraphs: ParagraphModel[] = [];
  const tables: TableModel[] = [];
  const sections: SectionModel[] = [];

  const buildParagraph = (
    p: XElement,
    index: number,
    blockIndex: number,
    insideTable: boolean
  ): ParagraphModel => {
    const pPr = childW(p, "pPr");
    const pStyleEl = childW(pPr, "pStyle");
    const styleId = pStyleEl ? getAttrW(pStyleEl, "val") : null;
    const style = styleId ? styles.get(styleId) : undefined;
    const chain = resolveStyleChain(styleId, styles);
    const direct = parseParaProps(pPr);

    let effP: ParaProps = { ...docDefaults.pPr };
    let styleR: RunProps = { ...docDefaults.rPr };
    // Default paragraph style (usually "Normal") applies when no style is set.
    const defaultPara = [...styles.values()].find(
      (s) => s.type === "paragraph" && s.isDefault
    );
    if (!styleId && defaultPara) {
      effP = mergePara(effP, defaultPara.pPr);
      styleR = mergeRun(styleR, defaultPara.rPr);
    }
    for (const s of chain) {
      effP = mergePara(effP, s.pPr);
      styleR = mergeRun(styleR, s.rPr);
    }
    effP = mergePara(effP, direct);

    const runs: RunModel[] = [];
    for (const r of childrenW(p, "r")) {
      const rPr = childW(r, "rPr");
      const props = parseRunProps(rPr);
      let text = "";
      for (const t of childrenW(r, "t")) text += t.textContent ?? "";
      const hasPageBreak = childrenW(r, "br").some(
        (br) => getAttrW(br, "type") === "page"
      );
      const hasDrawing =
        r.getElementsByTagNameNS(NS.w, "drawing").length > 0 ||
        r.getElementsByTagName("w:pict").length > 0;
      runs.push({
        el: r,
        text,
        props,
        effective: mergeRun(styleR, props),
        hasPageBreak,
        hasDrawing,
      });
    }

    // Dominant run = the one carrying the most text.
    let dominant: RunModel | null = null;
    for (const r of runs) {
      if (!dominant || r.text.length > dominant.text.length) dominant = r;
    }
    const runProps = dominant ? dominant.effective : styleR;

    const text = paragraphText(p).replace(/ /g, " ");
    const numPr = childW(pPr, "numPr");
    return {
      index,
      blockIndex,
      el: p,
      styleId: styleId ?? null,
      styleName: style?.name ?? null,
      text,
      isEmpty: text.trim().length === 0 && !runs.some((r) => r.hasDrawing),
      direct,
      props: effP,
      runProps,
      runs,
      hasDrawing: runs.some((r) => r.hasDrawing),
      hasPageBreakAfterInRuns: runs.some((r) => r.hasPageBreak),
      hasNumbering: numPr != null,
      hasSectPr: childW(pPr, "sectPr") != null,
      insideTable,
    };
  };

  const parseSectPr = (sectPr: XElement): SectionModel => {
    const pgMar = childW(sectPr, "pgMar");
    const pgSz = childW(sectPr, "pgSz");
    const pgNumType = childW(sectPr, "pgNumType");
    const num = (el: XElement | null, attr: string): number | undefined => {
      const v = el ? getAttrW(el, attr) : null;
      return v != null ? Number.parseInt(v, 10) : undefined;
    };
    const headerRefs: { type: string; rId: string }[] = [];
    const footerRefs: { type: string; rId: string }[] = [];
    for (const ref of childrenW(sectPr, "headerReference")) {
      headerRefs.push({
        type: getAttrW(ref, "type") ?? "default",
        rId: ref.getAttributeNS(NS.r, "id") ?? "",
      });
    }
    for (const ref of childrenW(sectPr, "footerReference")) {
      footerRefs.push({
        type: getAttrW(ref, "type") ?? "default",
        rId: ref.getAttributeNS(NS.r, "id") ?? "",
      });
    }
    return {
      el: sectPr,
      marginTop: num(pgMar, "top"),
      marginBottom: num(pgMar, "bottom"),
      marginLeft: num(pgMar, "left"),
      marginRight: num(pgMar, "right"),
      pageWidth: num(pgSz, "w"),
      pageHeight: num(pgSz, "h"),
      headerRefs,
      footerRefs,
      titlePg: childW(sectPr, "titlePg") != null,
      pgNumStart: num(pgNumType, "start"),
    };
  };

  let pIndex = 0;
  let blockIndex = 0;
  for (let n = body.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as XElement;
    if (el.namespaceURI !== NS.w) continue;
    if (el.localName === "p") {
      const pm = buildParagraph(el, pIndex, blockIndex, false);
      paragraphs.push(pm);
      const pPr = childW(el, "pPr");
      const sectPr = childW(pPr, "sectPr");
      if (sectPr) sections.push(parseSectPr(sectPr));
      pIndex++;
      blockIndex++;
    } else if (el.localName === "tbl") {
      const rows = childrenW(el, "tr").length;
      const firstRow = childrenW(el, "tr")[0];
      const cols = firstRow ? childrenW(firstRow, "tc").length : 0;
      const firstCell = firstRow ? childrenW(firstRow, "tc")[0] : undefined;
      let firstCellText = "";
      if (firstCell) {
        for (const cp of childrenW(firstCell, "p")) {
          firstCellText += paragraphText(cp) + " ";
        }
      }
      tables.push({
        index: tables.length,
        blockIndex,
        el,
        rows,
        cols,
        firstCellText: firstCellText.trim(),
      });
      blockIndex++;
    } else if (el.localName === "sectPr") {
      sections.push(parseSectPr(el));
    }
  }

  // Headers / footers with relationship resolution
  const headersFooters: HeaderFooterModel[] = [];
  const relsDoc = await pkg.getXmlIfPresent("word/_rels/document.xml.rels");
  const relTargets = new Map<string, string>();
  if (relsDoc?.documentElement) {
    const rels = relsDoc.documentElement.getElementsByTagNameNS(
      NS.rels,
      "Relationship"
    );
    for (let i = 0; i < rels.length; i++) {
      const rel = rels.item(i) as XElement;
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      if (id && target) relTargets.set(id, target);
    }
  }
  for (const section of sections) {
    for (const kind of ["header", "footer"] as const) {
      const refs = kind === "header" ? section.headerRefs : section.footerRefs;
      for (const ref of refs) {
        const target = relTargets.get(ref.rId);
        if (!target) continue;
        const partName = target.startsWith("/")
          ? target.slice(1)
          : `word/${target}`;
        if (headersFooters.some((h) => h.partName === partName && h.refType === ref.type)) {
          continue;
        }
        const xml = await pkg.getXmlIfPresent(partName);
        if (!xml?.documentElement) continue;
        const root = xml.documentElement as XElement;
        let text = "";
        for (const p of descendantsW(root, "p")) text += paragraphText(p) + " ";
        const instrs = fieldInstructions(root);
        headersFooters.push({
          partName,
          kind,
          refType: ref.type,
          text: text.trim(),
          hasPageField: instrs.some((s) => /\bPAGE\b/i.test(s)),
        });
      }
    }
  }

  const footnotesXml = await pkg.getXmlIfPresent("word/footnotes.xml");
  const endnotesXml = await pkg.getXmlIfPresent("word/endnotes.xml");
  const countNotes = (doc: XDocument | null, local: string): number => {
    if (!doc?.documentElement) return 0;
    const notes = childrenW(doc.documentElement as XElement, local);
    // exclude separator/continuation pseudo-notes (negative / 0 ids)
    return notes.filter((n) => {
      const id = getAttrW(n, "id");
      return id != null && Number.parseInt(id, 10) > 0;
    }).length;
  };

  const docEl = documentXml.documentElement as XElement;
  const imageCount =
    docEl.getElementsByTagNameNS(NS.w, "drawing").length +
    docEl.getElementsByTagName("w:pict").length;
  const hyperlinkCount = docEl.getElementsByTagNameNS(NS.w, "hyperlink").length;
  const equationCount =
    docEl.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/math",
      "oMath"
    ).length;

  return {
    pkg,
    documentXml,
    stylesXml,
    body,
    paragraphs,
    tables,
    sections,
    styles,
    styleIdByName: byName,
    docDefaults,
    headersFooters,
    footnoteCount: countNotes(footnotesXml, "footnote"),
    endnoteCount: countNotes(endnotesXml, "endnote"),
    imageCount,
    hyperlinkCount,
    equationCount,
  };
}

/**
 * Semantic content fingerprint used by content-preservation checks: the
 * document's visible text, image count, table shapes, hyperlinks, notes.
 */
export interface ContentFingerprint {
  text: string;
  /** Trimmed text of every non-empty paragraph, in document order. */
  paragraphs: string[];
  imageCount: number;
  tableShapes: string[];
  hyperlinkCount: number;
  footnoteCount: number;
  endnoteCount: number;
  equationCount: number;
}

export function contentFingerprint(model: DocumentModel): ContentFingerprint {
  const paragraphs = model.paragraphs
    .map((p) => p.text.trim())
    .filter((t) => t.length > 0);
  return {
    text: paragraphs.join("\n"),
    paragraphs,
    imageCount: model.imageCount,
    tableShapes: model.tables.map((t) => `${t.rows}x${t.cols}`),
    hyperlinkCount: model.hyperlinkCount,
    footnoteCount: model.footnoteCount,
    endnoteCount: model.endnoteCount,
    equationCount: model.equationCount,
  };
}
