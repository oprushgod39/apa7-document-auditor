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
  /** Explicit author marker such as "[H2]" or "Subheading 2:". */
  marker?: {
    raw: string;
    cleanText: string;
    kind: "apa_level" | "subheading" | "generic_subheading";
    ordinal?: number;
  };
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
  let genericSubheadingOrdinal = 0;

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i]!;
    if (p.isEmpty || p.insideTable) continue;
    // Skip title page and references entries; the References heading itself
    // is handled by the references rules (it is a Level-1-formatted heading).
    if (i < opts.bodyStartIndex) continue;
    if (opts.referencesHeadingIndex != null && i >= opts.referencesHeadingIndex) continue;

    const cls = classifyParagraph(p, paras, i);
    if (!cls) continue;
    if (cls.marker?.kind === "generic_subheading") {
      genericSubheadingOrdinal = Math.min(4, genericSubheadingOrdinal + 1);
      cls.marker.ordinal = genericSubheadingOrdinal;
      cls.level = genericSubheadingOrdinal + 1;
      cls.signals = [`generic subheading occurrence ${genericSubheadingOrdinal} → APA Level ${cls.level}`];
    } else if (cls.marker?.kind === "subheading") {
      genericSubheadingOrdinal = cls.marker.ordinal ?? 0;
    } else if (cls.level === 1) {
      // A new centered main heading starts a fresh subheading sequence.
      genericSubheadingOrdinal = 0;
    }
    out.push(cls);
  }
  return out;
}

/**
 * A source document may style its only tier of section headings as Word
 * "Heading 2" (or deeper) without ever using "Heading 1" — e.g. an author
 * who never opened the Styles pane past the second entry. Taken literally,
 * that produces left-aligned Level 2 headings throughout with no Level 1
 * anywhere, which is not what APA 7 hierarchy means: a document's *only*
 * tier of section headings is its top tier and belongs at centered Level 1,
 * regardless of which named Word style the author happened to pick.
 *
 * This only touches headings whose level came from a raw Word style
 * (`fromStyle`) with no author marker — an explicit `[H2]` or
 * `Subheading 1:` marker is deliberate intent and is never renumbered.
 * Heuristic (unstyled) headings already default to Level 1 by product
 * policy, so they're excluded from the "lowest level present" calculation.
 *
 * Must be called on the *final* heading list — after any caller-side
 * filtering (e.g. dropping a repeated document-title paragraph that
 * classified as a heading) — not inside `classifyHeadings` itself, since a
 * pseudo-heading like that repeated title can masquerade as a genuine
 * Level 1 and wrongly suppress normalization of the real section headings.
 */
export function normalizeStyleLevels(headings: ClassifiedHeading[]): void {
  const styleLevels = headings.filter((h) => h.fromStyle && !h.marker).map((h) => h.level);
  if (styleLevels.length === 0) return;
  const minLevel = Math.min(...styleLevels);
  if (minLevel <= 1) return;
  const offset = minLevel - 1;
  for (const h of headings) {
    if (!h.fromStyle || h.marker) continue;
    const original = h.level;
    h.level = Math.max(1, h.level - offset);
    h.signals.push(`normalized: no Level 1 heading in source, so Word style level ${original} → APA Level ${h.level}`);
  }
}

export function classifyParagraph(
  p: ParagraphModel,
  paras: ParagraphModel[],
  i: number
): ClassifiedHeading | null {
  const text = p.text.trim();
  if (text.length === 0) return null;
  // APA table/figure labels and notes are captions, never section headings.
  if (/^(?:table|figure)\s+\d+\.?$/i.test(text) || /^note\.\s/i.test(text)) return null;

  // Explicit author instructions are deterministic and take precedence over
  // source styles and visual heuristics. Supported examples:
  // Direct APA-level markers: [H2], [Heading 3], Heading 1:, Level 2:.
  const explicit = /^(\[(?:h|heading|level)\s*([1-5])\]|(?:heading|level)\s*([1-5])\s*(?::|-|–))\s*(.+)$/i.exec(text);
  if (explicit) {
    const level = Number(explicit[2] ?? explicit[3]);
    const cleanText = explicit[4]!.trim();
    if (cleanText) {
      return {
        paragraphIndex: p.index,
        text: cleanText,
        level,
        confidence: "high",
        score: 120,
        signals: [`explicit Level ${level} marker`],
        fromStyle: false,
        marker: {
          raw: text.slice(0, text.length - cleanText.length).trim(),
          cleanText,
          kind: "apa_level",
        },
      };
    }
  }
  // Product-facing subheading numbers are relative to the centered main
  // heading: Subheading 1 → APA Level 2, ... Subheading 4 → APA Level 5.
  const numberedSubheading = /^(\[?sub[\s-]*heading\s*([1-4])\]?\s*(?::|-|–)?\s*)(.+)$/i.exec(text);
  if (numberedSubheading && numberedSubheading[3]!.trim()) {
    const ordinal = Number(numberedSubheading[2]);
    const cleanText = numberedSubheading[3]!.trim();
    return {
      paragraphIndex: p.index,
      text: cleanText,
      level: ordinal + 1,
      confidence: "high",
      score: 120,
      signals: [`explicit Subheading ${ordinal} marker → APA Level ${ordinal + 1}`],
      fromStyle: false,
      marker: {
        raw: numberedSubheading[1]!.trim(),
        cleanText,
        kind: "subheading",
        ordinal,
      },
    };
  }
  const genericSubheading = /^(\[?sub[\s-]*heading\]?\s*(?::|-|–)?\s*)(.+)$/i.exec(text);
  if (genericSubheading && genericSubheading[2]!.trim()) {
    return {
      paragraphIndex: p.index,
      text: genericSubheading[2]!.trim(),
      level: 2,
      confidence: "high",
      score: 115,
      signals: ["generic subheading marker"],
      fromStyle: false,
      marker: {
        raw: genericSubheading[1]!.trim(),
        cleanText: genericSubheading[2]!.trim(),
        kind: "generic_subheading",
      },
    };
  }

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
  // A literal full stop is the highest-risk terminal punctuation for a false
  // positive — ordinary declarative sentences end in one far more often than
  // headings do — so it still requires bold (or centered/a section word) to
  // be considered at all. "?" and "!" are common on genuine unformatted
  // question-style or exclamatory subheadings ("How Strongly Do Consumers
  // React?") and are treated as heading-friendly, not risky, terminal marks.
  const endsWithFullStop = /\.$/.test(text) && !/^\d/.test(text);
  const endsWithQuestionOrBang = /[?!]$/.test(text);
  const hasCitation = /\((?:1[6-9]|20)\d{2}[a-z]?\)|\([^)]*,\s*(?:1[6-9]|20)\d{2}/.test(text);
  const sectionWord = SECTION_HEADING_WORDS.test(text);

  // Long paragraphs, list items, and citation-bearing text are body text.
  if (words > 14 || p.hasNumbering || hasCitation) return null;
  // Without a strong direct signal (bold, centered, a known section word), a
  // candidate is only considered when it doesn't end in a literal full stop
  // — i.e. it may still be an unformatted heading with no terminal
  // punctuation at all, or one ending in "?"/"!".
  if (!bold && !centered && !sectionWord && endsWithFullStop) return null;
  if (endsWithFullStop && !bold) return null;

  let score = 0;
  if (bold) { score += 30; signals.push("bold"); }
  if (centered) { score += 15; signals.push("centered"); }
  if (italic) { score += 5; signals.push("italic"); }
  if (words <= 8) { score += 15; signals.push("short"); }
  if (!endsWithFullStop) { score += 10; signals.push(endsWithQuestionOrBang ? "question/exclamatory heading" : "no terminal punctuation"); }
  if (isTitleCaseish(text)) { score += 10; signals.push("title case"); }
  if (sectionWord) { score += 25; signals.push("common section heading word"); }

  // Surrounding context: body paragraph following strengthens heading-ness.
  const next = paras.slice(i + 1).find((q) => !q.isEmpty);
  if (next && next.text.trim().split(/\s+/).length > 20) {
    score += 10;
    signals.push("followed by body paragraph");
  }

  if (score < 45) return null;

  // Per product policy, an unlabelled paragraph that is confidently a
  // heading defaults to APA Level 1. Authors can force Levels 2–5 with an
  // explicit marker, avoiding guesses based on source-document styling.
  const level = 1;

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
