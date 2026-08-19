import type { DocumentModel, ParagraphModel } from "../../docx/model.js";

/**
 * Multi-signal heading classifier.
 *
 * Signals: existing Word style, bold/italic weight, alignment, paragraph
 * length, terminal punctuation, capitalization, surrounding paragraphs,
 * and document position. Produces a level (1–5) or 0 (body) plus confidence.
 */

export interface ClassifiedHeading {
  paragraphIndex: number;
  text: string;
  level: number; // 1..5; 0 means "not a heading"
  confidence: "high" | "medium" | "low";
  score: number;
  signals: string[];
  fromStyle: boolean;
}

const STYLE_LEVEL: Record<string, number> = {
  heading1: 1, "heading 1": 1,
  heading2: 2, "heading 2": 2,
  heading3: 3, "heading 3": 3,
  heading4: 4, "heading 4": 4,
  heading5: 5, "heading 5": 5,
};

function isTitleCaseish(text: string): boolean {
  const words = text.split(/\s+/).filter((w) => w.length > 3);
  if (words.length === 0) return true;
  const capped = words.filter((w) => /^[A-ZÀ-Þ]/.test(w)).length;
  return capped / words.length >= 0.6;
}

const SECTION_HEADING_WORDS =
  /^(introduction|method|methods|methodology|results|discussion|conclusion|conclusions|literature review|background|limitations|implications|recommendations|summary|findings|analysis|participants|procedure|procedures|materials|measures|design|data analysis|future research|theoretical framework)$/i;

export function classifyHeadings(
  model: DocumentModel,
  opts: { bodyStartIndex: number; referencesHeadingIndex: number | null }
): ClassifiedHeading[] {
  const out: ClassifiedHeading[] = [];
  const paras = model.paragraphs;

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i]!;
    if (p.isEmpty || p.insideTable) continue;
    // Skip title page and references entries; the References heading itself
    // is handled by the references rules (it is a Level-1-formatted heading).
    if (i < opts.bodyStartIndex) continue;
    if (opts.referencesHeadingIndex != null && i >= opts.referencesHeadingIndex) continue;

    const cls = classifyParagraph(p, paras, i);
    if (cls) out.push(cls);
  }
  return out;
}

export function classifyParagraph(
  p: ParagraphModel,
  paras: ParagraphModel[],
  i: number
): ClassifiedHeading | null {
  const text = p.text.trim();
  if (text.length === 0) return null;

  const signals: string[] = [];

  // --- Signal 1: explicit Word heading style (dominant) ----------------
  const styleKey = (p.styleName ?? p.styleId ?? "").toLowerCase();
  const styleLevel = STYLE_LEVEL[styleKey];
  if (styleLevel) {
    return {
      paragraphIndex: p.index,
      text,
      level: styleLevel,
      confidence: "high",
      score: 100,
      signals: [`Word style "${p.styleName ?? p.styleId}"`],
      fromStyle: true,
    };
  }

  // --- Heuristic scoring for unstyled paragraphs -----------------------
  const words = text.split(/\s+/).length;
  const bold = p.runProps.bold === true;
  const italic = p.runProps.italic === true;
  const centered = p.props.alignment === "center";
  const endsWithPeriod = /[.!?]$/.test(text) && !/^\d/.test(text);
  const hasCitation = /\((?:1[6-9]|20)\d{2}[a-z]?\)|\([^)]*,\s*(?:1[6-9]|20)\d{2}/.test(text);
  const sectionWord = SECTION_HEADING_WORDS.test(text);

  // Long paragraphs, list items, and citation-bearing text are body text.
  if (words > 14 || p.hasNumbering || hasCitation) return null;
  if (!bold && !centered && !sectionWord) return null;
  if (endsWithPeriod && !bold) return null;

  let score = 0;
  if (bold) { score += 30; signals.push("bold"); }
  if (centered) { score += 15; signals.push("centered"); }
  if (italic) { score += 5; signals.push("italic"); }
  if (words <= 8) { score += 15; signals.push("short"); }
  if (!endsWithPeriod) { score += 10; signals.push("no terminal period"); }
  if (isTitleCaseish(text)) { score += 10; signals.push("title case"); }
  if (sectionWord) { score += 25; signals.push("common section heading word"); }

  // Surrounding context: body paragraph following strengthens heading-ness.
  const next = paras.slice(i + 1).find((q) => !q.isEmpty);
  if (next && next.text.trim().split(/\s+/).length > 20) {
    score += 10;
    signals.push("followed by body paragraph");
  }

  if (score < 45) return null;

  // Level inference for unstyled headings.
  let level: number;
  if (centered) level = 1;
  else if (bold && italic) level = 3;
  else if (bold) level = 2;
  else level = 1; // sectionWord uncentered, unbolded

  const confidence: "high" | "medium" | "low" =
    score >= 75 ? "high" : score >= 55 ? "medium" : "low";

  return {
    paragraphIndex: p.index,
    text,
    level,
    confidence,
    score,
    signals,
    fromStyle: false,
  };
}

/** Validate heading hierarchy; returns indexes of suspicious level skips. */
export function findHierarchySkips(
  headings: ClassifiedHeading[]
): { heading: ClassifiedHeading; previousLevel: number }[] {
  const out: { heading: ClassifiedHeading; previousLevel: number }[] = [];
  let prev = 0;
  for (const h of headings) {
    if (h.level > prev + 1 && prev > 0) {
      out.push({ heading: h, previousLevel: prev });
    }
    prev = h.level;
  }
  return out;
}
