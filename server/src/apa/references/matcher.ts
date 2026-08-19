import type { ParsedCitation } from "../citations/parser.js";
import type { ParsedReference } from "./parser.js";

/**
 * Bidirectional citation ↔ reference matching.
 *
 * Handles: et al., organization authors, punctuation variants, initials,
 * same-author same-year suffixes (2024a/2024b), minor normalization.
 */

export type MatchStatus =
  | "exact"
  | "probable"
  | "missing_reference"
  | "ambiguous";

export interface CitationMatch {
  citationIndex: number;
  status: MatchStatus;
  referenceIndexes: number[]; // candidate reference paragraph-list indexes
  detail: string;
}

export interface ReferenceUsage {
  referenceIndex: number;
  cited: boolean;
  citationIndexes: number[];
}

export interface MatchingResult {
  citationMatches: CitationMatch[];
  referenceUsage: ReferenceUsage[];
}

const ORG_ABBREVIATIONS: Record<string, string[]> = {
  // common expansions so "(APA, 2020)" matches "American Psychological Association"
  apa: ["american psychological association"],
  who: ["world health organization"],
  cdc: ["centers for disease control and prevention"],
  nih: ["national institutes of health"],
  un: ["united nations"],
  unesco: ["united nations educational, scientific and cultural organization"],
};

function normName(s: string): string {
  return s
    .toLowerCase()
    .replace(/[’']/g, "'")
    .replace(/\./g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function yearBase(year: string): string {
  return year.replace(/[a-z]$/, "");
}

function surnameMatches(citName: string, refSurname: string): boolean {
  const a = normName(citName);
  const b = normName(refSurname);
  if (a === b) return true;
  // Organization abbreviation vs full name
  const expansions = ORG_ABBREVIATIONS[a];
  if (expansions?.some((e) => e === b || b.includes(e))) return true;
  // Citation may use full org name; the reference stores full string.
  if (a.length > 8 && (b.includes(a) || a.includes(b))) return true;
  return false;
}

function citationMatchesReference(
  cit: ParsedCitation,
  ref: ParsedReference
): "exact" | "probable" | null {
  if (ref.surnames.length === 0) return null;

  // Year comparison. Suffixed years must match exactly; otherwise base match.
  let yearOk: "exact" | "probable" | null = null;
  if (ref.year == null) {
    yearOk = null;
  } else if (cit.year === ref.year) {
    yearOk = "exact";
  } else if (yearBase(cit.year) === yearBase(ref.year)) {
    // e.g. citation "2024" vs reference "2024a" — plausible but imprecise
    yearOk = "probable";
  }
  if (!yearOk) return null;

  const first = cit.authors[0];
  if (!first) return null;
  if (!surnameMatches(first, ref.surnames[0]!)) return null;

  if (cit.etAl) {
    // et al. requires the reference to plausibly have 3+ authors, but we
    // accept 2+ as probable (users sometimes misuse et al.)
    return ref.surnames.length >= 3 && yearOk === "exact" ? "exact" : "probable";
  }
  if (cit.authors.length === 1) {
    if (ref.isOrganizationAuthor || ref.surnames.length === 1) {
      return yearOk === "exact" ? "exact" : "probable";
    }
    return "probable"; // single-name citation of multi-author work
  }
  // Two+ author citation: all cited surnames must appear in order.
  const allMatch = cit.authors.every((a, i) => {
    const rs = ref.surnames[i];
    return rs != null && surnameMatches(a, rs);
  });
  if (allMatch && cit.authors.length === ref.surnames.length) {
    return yearOk === "exact" ? "exact" : "probable";
  }
  if (allMatch) return "probable";
  return null;
}

export function matchCitationsToReferences(
  citations: ParsedCitation[],
  references: ParsedReference[]
): MatchingResult {
  const citationMatches: CitationMatch[] = [];
  const usage: ReferenceUsage[] = references.map((_, i) => ({
    referenceIndex: i,
    cited: false,
    citationIndexes: [],
  }));

  citations.forEach((cit, ci) => {
    const exact: number[] = [];
    const probable: number[] = [];
    references.forEach((ref, ri) => {
      const res = citationMatchesReference(cit, ref);
      if (res === "exact") exact.push(ri);
      else if (res === "probable") probable.push(ri);
    });

    let status: MatchStatus;
    let chosen: number[];
    if (exact.length === 1) {
      status = "exact";
      chosen = exact;
    } else if (exact.length > 1) {
      status = "ambiguous";
      chosen = exact;
    } else if (probable.length === 1) {
      status = "probable";
      chosen = probable;
    } else if (probable.length > 1) {
      status = "ambiguous";
      chosen = probable;
    } else {
      status = "missing_reference";
      chosen = [];
    }
    citationMatches.push({
      citationIndex: ci,
      status,
      referenceIndexes: chosen,
      detail:
        status === "missing_reference"
          ? `No reference-list entry found for ${cit.raw}`
          : status === "ambiguous"
            ? `${chosen.length} reference entries could match ${cit.raw}`
            : "",
    });
    for (const ri of chosen) {
      usage[ri]!.cited = true;
      usage[ri]!.citationIndexes.push(ci);
    }
  });

  return { citationMatches, referenceUsage: usage };
}

/** Sort key for APA alphabetical reference order. */
export function referenceSortKey(ref: ParsedReference): string {
  const author = normName(ref.authorsRaw || ref.raw).replace(/[^a-z0-9 ]/g, "");
  const year = ref.year ?? "";
  return `${author}|${year}`;
}
