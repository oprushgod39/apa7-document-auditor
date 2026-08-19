import { NS, type XElement } from "./xml.js";

/**
 * Cross-run-safe text replacement.
 *
 * A paragraph's visible text is spread across multiple w:t nodes. This
 * gathers them with offsets, finds the target string in the concatenation,
 * and edits only the affected text nodes — never touching run properties,
 * hyperlinks, images, or field codes.
 */

interface TextNodeRef {
  el: XElement; // the w:t element
  start: number;
  end: number;
}

function collectTextNodes(p: XElement): { nodes: TextNodeRef[]; full: string } {
  const nodes: TextNodeRef[] = [];
  let full = "";
  const walk = (node: XElement) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const e = n as XElement;
      if (e.namespaceURI === NS.w && e.localName === "t") {
        const text = e.textContent ?? "";
        nodes.push({ el: e, start: full.length, end: full.length + text.length });
        full += text;
        continue;
      }
      if (e.namespaceURI === NS.w && (e.localName === "instrText" || e.localName === "delText")) {
        continue;
      }
      walk(e);
    }
  };
  walk(p);
  return { nodes, full };
}

function setTText(el: XElement, text: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  if (text.length > 0) {
    el.appendChild(el.ownerDocument!.createTextNode(text));
  }
  if (/^\s|\s$/.test(text) || text.length === 0) {
    el.setAttribute("xml:space", "preserve");
  }
}

/**
 * Replace the first occurrence of `before` with `after` in the paragraph's
 * text. Returns true when a replacement was made.
 */
export function replaceParagraphText(
  p: XElement,
  before: string,
  after: string
): boolean {
  const { nodes, full } = collectTextNodes(p);
  const idx = full.indexOf(before);
  if (idx < 0) return false;
  const endIdx = idx + before.length;

  let replaced = false;
  for (const node of nodes) {
    if (node.end <= idx || node.start >= endIdx) continue; // untouched
    const text = node.el.textContent ?? "";
    const localStart = Math.max(0, idx - node.start);
    const localEnd = Math.min(text.length, endIdx - node.start);
    const prefix = text.slice(0, localStart);
    const suffix = text.slice(localEnd);
    if (!replaced) {
      setTText(node.el, prefix + after + suffix);
      replaced = true;
    } else {
      setTText(node.el, prefix + suffix);
    }
  }
  return replaced;
}

/** Full concatenated w:t text of a paragraph (no tab/break padding). */
export function rawParagraphText(p: XElement): string {
  return collectTextNodes(p).full;
}

/**
 * Strip leading whitespace (spaces/tabs typed as text, and w:tab run
 * children) from the start of a paragraph. Used when manual whitespace was
 * used to fake a first-line indent. Returns description of what was removed,
 * or null if nothing to strip.
 */
export function stripLeadingWhitespace(p: XElement): string | null {
  const removed: string[] = [];
  // Remove leading w:tab elements inside the first runs.
  outer: for (let n = p.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as XElement;
    if (e.namespaceURI !== NS.w) continue;
    if (e.localName === "pPr") continue;
    if (e.localName !== "r") break;
    // Within the run: drop leading tabs, then trim leading spaces of first w:t.
    for (let c = e.firstChild; c; ) {
      const next = c.nextSibling;
      if (c.nodeType === 1) {
        const ce = c as XElement;
        if (ce.namespaceURI === NS.w && ce.localName === "tab") {
          e.removeChild(ce);
          removed.push("tab");
          c = next;
          continue;
        }
        if (ce.namespaceURI === NS.w && ce.localName === "rPr") {
          c = next;
          continue;
        }
        if (ce.namespaceURI === NS.w && ce.localName === "t") {
          const text = ce.textContent ?? "";
          const trimmed = text.replace(/^[\s\t]+/, "");
          if (trimmed !== text) {
            removed.push(`${text.length - trimmed.length} space(s)`);
            setTText(ce, trimmed);
          }
          if (trimmed.length > 0) break outer; // real text reached
          c = next;
          continue;
        }
        break outer; // any other content (drawing, break) ends stripping
      }
      c = next;
    }
  }
  return removed.length > 0 ? removed.join(", ") : null;
}
