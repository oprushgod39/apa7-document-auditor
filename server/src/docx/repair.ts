import type { DocxPackage } from "./package.js";
import type { ParagraphModel } from "./model.js";
import { NS, childrenW, createW, getAttrW, type XDocument, type XElement } from "./xml.js";

/**
 * Structural document-repair helpers. Unlike docx/edit.ts (which changes
 * formatting properties on existing paragraphs) these functions restructure
 * the paragraph tree itself — currently just one case: a "References" (or
 * equivalent) heading label that was typed as a run at the tail of a body
 * paragraph, separated only by a manual line break (Shift+Enter) rather than
 * living in its own paragraph. Word treats `<w:br/>` (no `type`, or
 * `type="textWrapping"`) as a soft line break within one paragraph, so this
 * never becomes a real paragraph boundary the rest of the analyzer can see.
 *
 * Shared with server/src/apa/analysis.ts, which is why the heading regex and
 * the "looks like a reference entry" heuristic live here rather than there:
 * both modules need them and analysis.ts already depends on docx/*.
 */

/** Matches a References-list heading, and nothing else, in a trimmed string. */
export const REFERENCES_HEADINGS =
  /^(references|reference list|works cited|bibliography|reference)$/i;

/** Same heuristic analysis.ts uses to confirm a heading is followed by real entries. */
export function looksLikeReferenceEntry(text: string): boolean {
  return /\((?:1[6-9]|20)\d{2}[a-z]?\)|\(n\.d\.\)|https?:\/\//.test(text);
}

/** Concatenated text of a set of sibling paragraph-content nodes (runs, hyperlinks, ...). */
function textOfNode(node: XElement): string {
  let out = "";
  const walk = (el: XElement) => {
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const e = n as XElement;
      if (e.namespaceURI === NS.w) {
        if (e.localName === "t") {
          out += e.textContent ?? "";
          continue;
        }
        if (e.localName === "tab" || e.localName === "br") {
          out += " ";
          continue;
        }
        if (e.localName === "delText" || e.localName === "instrText") continue;
      }
      walk(e);
    }
  };
  walk(node);
  return out;
}

/** True for a manual line break (Shift+Enter) — not a page or column break. */
function isPlainLineBreak(br: XElement): boolean {
  const type = getAttrW(br, "type");
  return type == null || type === "textWrapping";
}

export interface DetectedEmbeddedHeading {
  /** The `<w:r>` that contains the qualifying line break; discarded on split. */
  breakRunEl: XElement;
  /** Trimmed text that remains in the paragraph before the break. */
  beforeText: string;
  /** Trimmed text after the break — matched the heading pattern. */
  headingText: string;
}

/**
 * Look for a heading label fused onto the tail of `p` via a plain line
 * break. Read-only — does not touch the DOM. Returns null unless every
 * condition holds:
 *  - there is at least one plain (`textWrapping`/untyped) `<w:br/>` inside a
 *    run that is a direct child of the paragraph (page/column breaks are
 *    ignored, and only the *last* qualifying break is considered);
 *  - the break's own run carries no other text (a break sharing a run with
 *    real text is a more unusual shape this repair does not attempt);
 *  - the text before the break is non-empty (a real preceding sentence);
 *  - the text after the break, trimmed, exactly matches `pattern` with
 *    nothing else around it.
 */
export function detectEmbeddedReferencesHeading(
  p: XElement,
  pattern: RegExp = REFERENCES_HEADINGS
): DetectedEmbeddedHeading | null {
  const contentChildren: XElement[] = [];
  for (let n = p.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as XElement;
    if (e.namespaceURI === NS.w && e.localName === "pPr") continue;
    contentChildren.push(e);
  }

  let breakRunIndex = -1;
  for (let i = 0; i < contentChildren.length; i++) {
    const c = contentChildren[i]!;
    if (c.namespaceURI !== NS.w || c.localName !== "r") continue;
    const hasQualifyingBreak = childrenW(c, "br").some(isPlainLineBreak);
    if (hasQualifyingBreak) breakRunIndex = i; // keep the last one found
  }
  if (breakRunIndex === -1) return null;

  const breakRun = contentChildren[breakRunIndex]!;
  // A break run that also carries its own text is a shape this repair does
  // not attempt to split mid-run; leave it alone.
  if (childrenW(breakRun, "t").length > 0) return null;

  const beforeNodes = contentChildren.slice(0, breakRunIndex);
  const afterNodes = contentChildren.slice(breakRunIndex + 1);
  if (afterNodes.length === 0) return null;

  const beforeText = beforeNodes.map(textOfNode).join("").trim();
  const headingText = afterNodes.map(textOfNode).join("").trim();
  if (!beforeText) return null;
  if (!pattern.test(headingText)) return null;

  return { breakRunEl: breakRun, beforeText, headingText };
}

/**
 * Search a document's paragraphs, from the end, for one paragraph whose tail
 * is a References-style heading fused on via a line break (see
 * `detectEmbeddedReferencesHeading`), immediately followed by paragraphs that
 * look like real reference entries. Mirrors the standalone-heading search in
 * analysis.ts. Read-only.
 */
export function findEmbeddedReferencesHeadingCandidate(
  paragraphs: ParagraphModel[]
): { paragraph: ParagraphModel; detected: DetectedEmbeddedHeading } | null {
  for (let i = paragraphs.length - 1; i >= 0; i--) {
    const p = paragraphs[i]!;
    if (p.insideTable) continue;
    const detected = detectEmbeddedReferencesHeading(p.el);
    if (!detected) continue;
    const following = paragraphs.slice(i + 1).filter((q) => !q.isEmpty);
    const looksLikeRefs = following.some((q) => looksLikeReferenceEntry(q.text));
    if (looksLikeRefs || following.length === 0) {
      return { paragraph: p, detected };
    }
  }
  return null;
}

/**
 * Physically split `p` at the detected break: everything from the break
 * run onward is removed from `p` and moved into a brand-new paragraph
 * inserted immediately after it in the same parent. The break run itself is
 * discarded (the new paragraph does not need a leading line break) — the
 * runs after it move over as-is, including their own `<w:rPr>`, so the
 * caller (e.g. APA-REFERENCE-001) can apply whatever heading formatting it
 * needs on top. The new paragraph's `<w:pPr>` starts absent.
 *
 * Marks `word/document.xml` dirty. Returns the new paragraph element.
 */
export function splitParagraphAtEmbeddedHeading(
  doc: XDocument,
  pkg: DocxPackage,
  p: XElement,
  detected: DetectedEmbeddedHeading
): XElement {
  const parent = p.parentNode as XElement;
  const newParagraph = createW(doc, "p");

  const toMove: NonNullable<typeof detected.breakRunEl.nextSibling>[] = [];
  for (let n = detected.breakRunEl.nextSibling; n; n = n.nextSibling) {
    toMove.push(n);
  }
  for (const n of toMove) {
    newParagraph.appendChild(n); // appendChild reparents the node, removing it from p
  }
  p.removeChild(detected.breakRunEl);

  if (p.nextSibling) parent.insertBefore(newParagraph, p.nextSibling);
  else parent.appendChild(newParagraph);

  pkg.markDirty("word/document.xml");
  return newParagraph;
}
