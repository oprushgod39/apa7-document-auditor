import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type { Document as XDocument, Element as XElement } from "@xmldom/xmldom";

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

export type { XDocument, XElement };

/**
 * Parse XML strictly. @xmldom/xmldom does not resolve external entities or
 * DTDs (no network / filesystem access), which protects against XXE by design.
 * We additionally reject any document containing a DOCTYPE declaration.
 */
export function parseXml(content: string, partName: string): XDocument {
  if (/<!DOCTYPE/i.test(content)) {
    throw new Error(`XML part ${partName} contains a DOCTYPE declaration`);
  }
  let fatal: string | null = null;
  const parser = new DOMParser({
    onError: (level, message) => {
      if (level === "fatalError") fatal = message;
    },
  });
  const doc = parser.parseFromString(content, "text/xml");
  if (fatal || !doc.documentElement) {
    throw new Error(`XML part ${partName} could not be parsed: ${fatal ?? "empty"}`);
  }
  return doc;
}

const serializer = new XMLSerializer();

export function serializeXml(doc: XDocument): string {
  const body = serializer.serializeToString(doc);
  if (body.startsWith("<?xml")) return body;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` + body;
}

/** First direct child with the given wordprocessingml local name, or null. */
export function childW(el: XElement | null | undefined, local: string): XElement | null {
  if (!el) return null;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (n as XElement).localName === local && (n as XElement).namespaceURI === NS.w) {
      return n as XElement;
    }
  }
  return null;
}

/** All direct children with the given wordprocessingml local name. */
export function childrenW(el: XElement | null | undefined, local: string): XElement[] {
  const out: XElement[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (n as XElement).localName === local && (n as XElement).namespaceURI === NS.w) {
      out.push(n as XElement);
    }
  }
  return out;
}

/** All descendants (document order) with the given w: local name. */
export function descendantsW(el: XElement | null | undefined, local: string): XElement[] {
  const out: XElement[] = [];
  if (!el) return out;
  const walk = (node: XElement) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const e = n as XElement;
      if (e.localName === local && e.namespaceURI === NS.w) out.push(e);
      walk(e);
    }
  };
  walk(el);
  return out;
}

export function getAttrW(el: XElement | null | undefined, name: string): string | null {
  if (!el) return null;
  const v = el.getAttributeNS(NS.w, name);
  return v === "" ? null : v;
}

export function setAttrW(el: XElement, name: string, value: string): void {
  el.setAttributeNS(NS.w, `w:${name}`, value);
}

export function removeAttrW(el: XElement, name: string): void {
  el.removeAttributeNS(NS.w, name);
}

/** Create a w: element within the given document. */
export function createW(doc: XDocument, local: string): XElement {
  return doc.createElementNS(NS.w, `w:${local}`);
}

/**
 * Get or create a direct w: child. OOXML property containers are
 * order-sensitive; `before` lists local names that must come after the new
 * element (insert before the first of them found).
 */
export function ensureChildW(
  doc: XDocument,
  parent: XElement,
  local: string,
  before: string[] = []
): XElement {
  const existing = childW(parent, local);
  if (existing) return existing;
  const el = createW(doc, local);
  let ref: XElement | null = null;
  for (const b of before) {
    const found = childW(parent, b);
    if (found) {
      ref = found;
      break;
    }
  }
  if (ref) parent.insertBefore(el, ref);
  else parent.appendChild(el);
  return el;
}

export function removeChildW(parent: XElement, local: string): boolean {
  const el = childW(parent, local);
  if (el) {
    parent.removeChild(el);
    return true;
  }
  return false;
}

/** Concatenated text of all w:t descendants (plus tabs/breaks as spaces). */
export function paragraphText(p: XElement): string {
  let out = "";
  const walk = (node: XElement) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
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
        // Skip deleted text and instruction text
        if (e.localName === "delText" || e.localName === "instrText") continue;
      }
      walk(e);
    }
  };
  walk(p);
  return out;
}
