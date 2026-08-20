import mammoth from "mammoth";
import { REFERENCE_HEADING } from "./word_count.js";

/**
 * Converts a source .docx to HTML with images inlined as base64 data URIs
 * (mammoth strips headers/footers by default, matching what
 * buildDocumentModel/countDocumentWords consider "body content").
 */
export async function docxToHtml(buffer: Buffer): Promise<string> {
  const result = await mammoth.convertToHtml(
    { buffer },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        const base64 = await image.read("base64");
        return { src: `data:${image.contentType};base64,${base64}` };
      }),
    }
  );
  return result.value;
}

export interface TopLevelBlock {
  /** Raw HTML for this single top-level element (e.g. a whole <p>...</p> or <table>...</table>). */
  html: string;
  /** Plain-text content of the element, tags stripped, entities decoded. */
  text: string;
}

const VOID_TAGS = new Set([
  "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr",
]);

const TAG_RE = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b[^>]*?(\/?)>/g;

const ENTITIES: Record<string, string> = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ",
};

function decodeEntities(s: string): string {
  return s.replace(/&(#\d+|#x[0-9a-fA-F]+|[a-zA-Z]+);/g, (full, code: string) => {
    if (code[0] === "#") {
      const cp = code[1] === "x" || code[1] === "X" ? Number.parseInt(code.slice(2), 16) : Number.parseInt(code.slice(1), 10);
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : full;
    }
    return ENTITIES[code] ?? full;
  });
}

function stripTags(html: string): string {
  return decodeEntities(html.replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
}

/**
 * Splits mammoth's flat HTML output into its top-level elements, one entry
 * per source document body block (paragraph, heading, table, or list),
 * preserving document order. Mammoth emits one top-level element per body
 * block for plain paragraphs/tables; consecutive list items get grouped
 * into a single <ul>/<ol>, which is still exactly one top-level block.
 */
export function splitTopLevelHtml(html: string): TopLevelBlock[] {
  const blocks: TopLevelBlock[] = [];
  let depth = 0;
  let blockStart = -1;
  TAG_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = TAG_RE.exec(html))) {
    const [full, closing, name, selfCloseMark] = match;
    const isVoid = VOID_TAGS.has(name!.toLowerCase()) || selfCloseMark === "/" || full.endsWith("/>");
    if (!closing) {
      if (depth === 0) blockStart = match.index;
      if (!isVoid) depth++;
    } else {
      depth = Math.max(0, depth - 1);
      if (depth === 0 && blockStart >= 0) {
        const end = match.index + full.length;
        const raw = html.slice(blockStart, end);
        blocks.push({ html: raw, text: stripTags(raw) });
        blockStart = -1;
      }
    }
  }
  return blocks;
}

/**
 * Converts a source document to HTML and strips everything from its
 * References/Bibliography/Works Cited heading onward (inclusive), mirroring
 * the exact heading detection used by countDocumentWords so the merged PDF
 * matches the word counts shown in the merge preview.
 *
 * Unlike the word-count path (which locates the cutoff by paragraph
 * blockIndex in the parsed OOXML model), this operates directly on the
 * rendered HTML's top-level elements and cuts at the first one whose text
 * matches the reference-heading pattern. That sidesteps having to assume
 * mammoth's HTML blocks stay 1:1 with document.xml's block indexes (an
 * assumption that breaks once a document contains bulleted/numbered lists,
 * which mammoth groups into a single <ul>/<ol> spanning several blockIndex
 * values) while still applying the identical regex to identical heading text.
 */
export async function docxToTrimmedHtml(buffer: Buffer): Promise<string> {
  const html = await docxToHtml(buffer);
  const blocks = splitTopLevelHtml(html);
  const cutoff = blocks.findIndex((block) => REFERENCE_HEADING.test(block.text));
  const kept = cutoff === -1 ? blocks : blocks.slice(0, cutoff);
  return kept.map((block) => block.html).join("");
}
