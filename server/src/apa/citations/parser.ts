/**
 * Deterministic APA 7 in-text citation parser.
 *
 * Recognizes parenthetical and narrative citations, single/two/3+ authors,
 * "et al.", organization authors, multiple citations in one parenthesis,
 * year suffixes (2024a), n.d., page/paragraph locators and ranges.
 */

export type CitationKind = "parenthetical" | "narrative";

export interface ParsedCitation {
  kind: CitationKind;
  raw: string;
  /** The full parenthetical group this citation came from (for context). */
  groupRaw: string;
  paragraphIndex: number;
  /** Author surnames / organization names as written. */
  authors: string[];
  etAl: boolean;
  year: string; // e.g. "2024", "2024a", "n.d."
  locator?: string; // "p. 21", "pp. 21–24", "para. 4"
  /** Problems detected in the citation mechanics. */
  problems: CitationProblem[];
}

export interface CitationProblem {
  code:
    | "AMPERSAND_IN_NARRATIVE"
    | "AND_IN_PARENTHETICAL"
    | "ET_AL_MISSING_PERIOD"
    | "MISSING_COMMA_BEFORE_YEAR"
    | "LOCATOR_FORMAT";
  message: string;
}

const YEAR = String.raw`(?:(?:1[6-9]|20)\d{2}[a-z]?|n\.d\.|in press)`;
const LOCATOR = String.raw`(?:pp?\.\s*[\divxlc]+(?:\s*[–\-—]\s*[\divxlc]+)?(?:\s*,\s*\d+)*|para\.?s?\.?\s*\d+(?:\s*[–\-—]\s*\d+)?|Chapter\s+\d+|Table\s+\d+|Figure\s+\d+|\d+:\d+)`;

// A surname token: capitalized word possibly hyphenated/apostrophized,
// optionally with particles (van, de, etc.)
const SURNAME = String.raw`(?:(?:van|von|de|del|der|di|la|le|al|bin|ter)\s)?[A-ZÀ-Þ][\w'’\-À-ÿ]+`;

/** Extract citations from a paragraph's plain text. */
export function parseCitations(
  text: string,
  paragraphIndex: number
): ParsedCitation[] {
  const out: ParsedCitation[] = [];
  out.push(...parseParenthetical(text, paragraphIndex));
  out.push(...parseNarrative(text, paragraphIndex));
  return out;
}

function normalizeDashes(s: string): string {
  return s.replace(/–|—/g, "-");
}

function parseParenthetical(
  text: string,
  paragraphIndex: number
): ParsedCitation[] {
  const out: ParsedCitation[] = [];
  const parenRe = /\(([^()]+)\)/g;
  let m: RegExpExecArray | null;
  while ((m = parenRe.exec(text)) !== null) {
    const inner = m[1]!;
    // Skip if preceded immediately by an author name — that is a narrative
    // citation year-group handled by parseNarrative.
    const before = text.slice(Math.max(0, m.index - 60), m.index);
    if (isNarrativeYearGroup(before, inner)) continue;
    if (!new RegExp(YEAR).test(inner)) continue;
    // Split multiple citations: (Smith, 2022; Williams, 2023)
    const segments = inner.split(/;\s*/);
    for (const seg of segments) {
      const parsed = parseCitationSegment(seg.trim());
      if (parsed) {
        out.push({
          kind: "parenthetical",
          raw: `(${seg.trim()})`,
          groupRaw: `(${inner})`,
          paragraphIndex,
          ...parsed,
        });
      }
    }
  }
  return out;
}

function isNarrativeYearGroup(before: string, inner: string): boolean {
  // Pure year (possibly with locator) directly after a capitalized word or
  // "et al." → narrative form "Smith (2024)".
  const yearOnly = new RegExp(
    `^${YEAR}(?:\\s*,\\s*${LOCATOR})?$`,
    "i"
  ).test(inner.trim());
  if (!yearOnly) return false;
  return /(?:[A-ZÀ-Þ][\w'’\-À-ÿ]+|et al\.?|[A-Z]{2,})\s*$/.test(before);
}

interface SegmentParse {
  authors: string[];
  etAl: boolean;
  year: string;
  locator?: string;
  problems: CitationProblem[];
}

function parseCitationSegment(seg: string): SegmentParse | null {
  const problems: CitationProblem[] = [];
  const norm = normalizeDashes(seg);
  // author-part, year [, locator]
  const re = new RegExp(
    String.raw`^(.*?)(,)?\s+(${YEAR})(?:\s*,\s*(${LOCATOR}))?\s*$`,
    "i"
  );
  const m = re.exec(norm);
  if (!m) return null;
  const authorPart = m[1]!.trim().replace(/,$/, "");
  const hadComma = m[2] != null;
  const year = m[3]!;
  const locator = m[4] ?? undefined;
  if (authorPart.length === 0) return null;
  // Reject things that clearly are not citations (e.g. "(see Figure 1, 2020)")
  if (/^(see|cf\.|e\.g\.|i\.e\.)/i.test(authorPart) && !/[A-ZÀ-Þ]/.test(authorPart.slice(4))) {
    return null;
  }
  // Author part must start with a capital letter (name or org).
  if (!/^[A-ZÀ-Þ"“]/.test(authorPart)) return null;

  let etAl = false;
  let authorsRaw = authorPart;
  const etAlMatch = /\bet\s+al\.?(?=$|,)/.exec(authorsRaw);
  if (etAlMatch) {
    etAl = true;
    if (!etAlMatch[0].endsWith(".")) {
      problems.push({
        code: "ET_AL_MISSING_PERIOD",
        message: `"et al" requires a period: "et al."`,
      });
    }
    authorsRaw = authorsRaw.slice(0, etAlMatch.index).replace(/,\s*$/, "").trim();
  }

  if (/\band\b/.test(authorsRaw)) {
    problems.push({
      code: "AND_IN_PARENTHETICAL",
      message: `Parenthetical citations use "&" rather than "and".`,
    });
  }
  if (!hadComma) {
    problems.push({
      code: "MISSING_COMMA_BEFORE_YEAR",
      message: "A comma is required between authors and year.",
    });
  }

  const authors = splitAuthors(authorsRaw);
  if (authors.length === 0) return null;
  return { authors, etAl, year: year.toLowerCase() === "n.d." ? "n.d." : year, locator, problems };
}

function splitAuthors(raw: string): string[] {
  const cleaned = raw.replace(/\s+/g, " ").trim();
  if (cleaned.length === 0) return [];
  // Organization author (contains lowercase words like "of", "for", or is a
  // multi-word capitalized phrase without separators)
  if (!/[&]|,|\band\b/.test(cleaned)) {
    return [cleaned];
  }
  return cleaned
    .split(/\s*(?:,|&|\band\b)\s*/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

function parseNarrative(text: string, paragraphIndex: number): ParsedCitation[] {
  const out: ParsedCitation[] = [];
  // Author name(s) directly followed by (year[, locator])
  const re = new RegExp(
    String.raw`(${SURNAME}(?:\s+(?:and|&)\s+${SURNAME})?(?:,\s+${SURNAME}\s*,?\s+(?:and|&)\s+${SURNAME})?(?:\s+et\s+al\.?)?|[A-Z][A-Za-z&.\s]{2,60}?)\s*\((${YEAR})(?:\s*,\s*(${LOCATOR}))?\)`,
    "g"
  );
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const authorPart = m[1]!.trim();
    const year = m[2]!;
    const locator = m[3] ?? undefined;
    if (!/^[A-ZÀ-Þ]/.test(authorPart)) continue;
    // Filter obvious non-citations: sentence fragments ending prepositions.
    if (/\b(the|a|an|in|of|on|for|with|from|by|to|at|as)$/i.test(authorPart)) continue;

    const problems: CitationProblem[] = [];
    let etAl = false;
    let authorsRaw = authorPart;
    const etAlMatch = /\bet\s+al\.?$/.exec(authorsRaw);
    if (etAlMatch) {
      etAl = true;
      if (!etAlMatch[0].endsWith(".")) {
        problems.push({
          code: "ET_AL_MISSING_PERIOD",
          message: `"et al" requires a period: "et al."`,
        });
      }
      authorsRaw = authorsRaw.slice(0, etAlMatch.index).replace(/,\s*$/, "").trim();
    }
    if (/\s&\s/.test(authorsRaw)) {
      problems.push({
        code: "AMPERSAND_IN_NARRATIVE",
        message: `Narrative citations use "and" rather than "&".`,
      });
    }
    const authors = splitAuthors(authorsRaw.replace(/\s+(?:and)\s+/g, " & "));
    if (authors.length === 0) continue;
    out.push({
      kind: "narrative",
      raw: m[0],
      groupRaw: m[0],
      paragraphIndex,
      authors,
      etAl,
      year: year.toLowerCase() === "n.d." ? "n.d." : year,
      locator,
      problems,
    });
  }
  return out;
}
