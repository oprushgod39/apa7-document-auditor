import { diffChars } from "diff";

const blocks = new Set(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "tr"]);
const excluded = new Set(["script", "style", "meta", "link", "img", "svg", "math", "iframe", "object", "embed", "input", "button", "textarea", "select"]);

/** Keep pasted text and basic emphasis only; never retain untrusted attributes. */
export function sanitizeCitationHtml(html: string): string {
  const parsed = new DOMParser().parseFromString(html, "text/html");
  const output = document.createElement("div");
  // Word and browsers may put emphasis in clipboard CSS classes instead of
  // inline spans. Read only emphasis properties; never mount pasted CSS.
  const emphasisRules: { selector: string; italic: boolean; bold: boolean; underline: boolean }[] = [];
  for (const style of parsed.querySelectorAll("style")) {
    for (const match of (style.textContent ?? "").matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
      const declaration = document.createElement("span").style;
      declaration.cssText = match[2]!;
      const italic = /^(italic|oblique)/i.test(declaration.fontStyle) || /\b(?:italic|oblique)\b/i.test(declaration.font);
      const bold = /^(bold|[6-9]00)$/.test(declaration.fontWeight);
      const underline = declaration.textDecoration.includes("underline") || declaration.textDecorationLine.includes("underline");
      if (!italic && !bold && !underline) continue;
      for (const selector of match[1]!.split(",")) {
        if (selector.trim().startsWith("@")) continue;
        emphasisRules.push({ selector: selector.trim(), italic, bold, underline });
      }
    }
  }
  const copy = (node: Node, parent: HTMLElement) => {
    if (node.nodeType === Node.TEXT_NODE) { parent.appendChild(document.createTextNode(node.textContent ?? "")); return; }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    const tag = el.localName.toLowerCase();
    if (excluded.has(tag)) return;
    const rules = emphasisRules.filter((rule) => { try { return el.matches(rule.selector); } catch { return false; } });
    let target = parent;
    if (tag === "br") { parent.appendChild(document.createElement("br")); return; }
    if (blocks.has(tag)) { target = document.createElement("div"); parent.appendChild(target); }
    const format = (name: string) => { const next = document.createElement(name); target.appendChild(next); target = next; };
    if (["b", "strong"].includes(tag) || /^(bold|[6-9]00)$/.test(el.style.fontWeight) || rules.some((r) => r.bold)) format("strong");
    if (["i", "em"].includes(tag) || el.classList.contains("italic") || /^(italic|oblique)/i.test(el.style.fontStyle) || /\b(?:italic|oblique)\b/i.test(el.style.font) || rules.some((r) => r.italic)) format("em");
    if (tag === "u" || el.style.textDecoration.includes("underline") || rules.some((r) => r.underline)) format("u");
    if (["sup", "sub"].includes(tag)) format(tag);
    for (const child of el.childNodes) copy(child, target);
  };
  for (const child of parsed.body.childNodes) copy(child, output);
  return output.innerHTML;
}

function textMap(root: HTMLElement) {
  let text = "";
  const nodes: { node: Text; start: number; end: number }[] = [];
  const newline = () => { if (text && !text.endsWith("\n")) text += "\n"; };
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const value = node.textContent ?? "";
      nodes.push({ node: node as Text, start: text.length, end: text.length + value.length });
      text += value; return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    if (el.localName === "br") { text += "\n"; return; }
    const block = blocks.has(el.localName) && el !== root;
    if (block) newline();
    for (const child of el.childNodes) walk(child);
    if (block) newline();
  };
  walk(root);
  return { text: text.replace(/\n$/, ""), nodes };
}

export function citationHtmlText(html: string): string {
  const root = document.createElement("div"); root.innerHTML = sanitizeCitationHtml(html);
  return textMap(root).text;
}

export function plainCitationHtml(text: string): string {
  const root = document.createElement("div"); root.textContent = text;
  return root.innerHTML;
}

export function applyReferenceItalics(html: string, spans: { start: number; end: number }[]): string {
  const root = document.createElement("div"); root.innerHTML = sanitizeCitationHtml(html);
  const before = textMap(root).text;
  for (const span of [...spans].sort((a, b) => b.start - a.start)) {
    if (span.start < 0 || span.end <= span.start || span.end > before.length) continue;
    // Split text nodes rather than moving whole ranges, so existing links of
    // emphasis and paragraph boundaries stay intact.
    for (const item of textMap(root).nodes.reverse()) {
      const start = Math.max(item.start, span.start);
      const end = Math.min(item.end, span.end);
      if (start >= end || item.node.parentElement?.closest("em, i")) continue;
      const middle = item.node.splitText(start - item.start);
      middle.splitText(end - start);
      const em = document.createElement("em");
      middle.parentNode!.insertBefore(em, middle); em.appendChild(middle);
    }
  }
  if (textMap(root).text !== before) throw new Error("Reference formatting changed the text unexpectedly.");
  return root.innerHTML;
}

/** Apply citation edits at text offsets while leaving emphasis on other words intact. */
export function correctedCitationHtml(source: string, corrected: string): string {
  const root = document.createElement("div"); root.innerHTML = sanitizeCitationHtml(source);
  const original = textMap(root).text;
  const diff = diffChars(original, corrected, { timeout: 200, maxEditLength: 10_000 });
  if (!diff) throw new Error("Formatting could not be preserved. Use Word upload for this document.");
  let offset = 0;
  const edits: { start: number; length: number; value: string }[] = [];
  for (const part of diff) {
    if (!part.added && !part.removed) { offset += part.value.length; continue; }
    const previous = edits.at(-1);
    if (part.removed) { edits.push({ start: offset, length: part.value.length, value: "" }); offset += part.value.length; }
    else if (previous && previous.start + previous.length === offset) previous.value += part.value;
    else edits.push({ start: offset, length: 0, value: part.value });
  }
  for (const edit of edits.reverse()) {
    const { nodes } = textMap(root);
    const start = nodes.find((n) => n.start <= edit.start && n.end >= edit.start);
    const end = nodes.find((n) => n.start <= edit.start + edit.length && n.end >= edit.start + edit.length);
    if (!start || !end) throw new Error("Could not preserve this text's formatting. Use Word upload instead.");
    // Citation corrections never remove paragraph boundaries.
    const range = document.createRange();
    range.setStart(start.node, edit.start - start.start);
    range.setEnd(end.node, edit.start + edit.length - end.start);
    range.deleteContents(); range.insertNode(document.createTextNode(edit.value));
  }
  if (textMap(root).text !== corrected) throw new Error("Formatted text validation failed. Use Word upload instead.");
  return root.innerHTML;
}
