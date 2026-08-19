/**
 * Deterministic APA 7 reference-list entry parser and type classifier.
 * Parses author block, year, title, and source info; never invents data.
 */

export type ReferenceType =
  | "journal_article"
  | "book"
  | "edited_book"
  | "book_chapter"
  | "webpage"
  | "government_publication"
  | "report"
  | "newspaper_article"
  | "magazine_article"
  | "conference_paper"
  | "dissertation"
  | "thesis"
  | "dataset"
  | "software"
  | "video"
  | "podcast"
  | "online_media"
  | "organization_publication"
  | "unknown";

export interface ParsedReference {
  paragraphIndex: number;
  raw: string;
  /** Author block exactly as written (before the year parenthesis). */
  authorsRaw: string;
  /** Extracted surnames, lowercase-normalized for matching. */
  surnames: string[];
  isOrganizationAuthor: boolean;
  year: string | null; // "2024", "2024a", "n.d." or null when missing
  title: string | null;
  doi: string | null;
  url: string | null;
  type: ReferenceType;
  typeConfidence: "high" | "medium" | "low";
  /** Journal/volume/pages info if detected. */
  journal?: { name: string | null; volume: string | null; issue: string | null; pages: string | null };
  problems: ReferenceProblem[];
}

export interface ReferenceProblem {
  code:
    | "MISSING_YEAR"
    | "DOI_OLD_FORMAT"
    | "DOI_AS_LABEL"
    | "URL_RETRIEVED_FROM"
    | "AMPERSAND_MISSING"
    | "TRAILING_PERIOD_AFTER_DOI"
    | "POSSIBLE_MALFORMED";
  message: string;
  /** For deterministic fixes: exact before/after text when known. */
  before?: string;
  after?: string;
}

const DOI_RE = /(?:https?:\/\/(?:dx\.)?doi\.org\/|doi:\s*|DOI:\s*)(10\.\d{4,9}\/[^\s]+)/i;
const URL_RE = /https?:\/\/[^\s]+/i;

export function parseReference(text: string, paragraphIndex: number): ParsedReference {
  const raw = text.trim().replace(/\s+/g, " ");
  const problems: ReferenceProblem[] = [];

  // --- Year -----------------------------------------------------------
  const yearMatch = /\(((?:1[6-9]|20)\d{2}[a-z]?|n\.d\.)(?:,\s*[^)]+)?\)/.exec(raw);
  const year = yearMatch ? yearMatch[1]! : null;
  if (!year) {
    problems.push({
      code: "MISSING_YEAR",
      message: "No publication year found in (Year) position.",
    });
  }

  // --- Authors --------------------------------------------------------
  const authorsRaw = yearMatch ? raw.slice(0, yearMatch.index).trim() : "";
  const { surnames, isOrg } = extractSurnames(authorsRaw);

  // multiple authors separated by comma but missing "&" before last
  if (!isOrg && surnames.length >= 2 && !/&/.test(authorsRaw) && !/\bet al\./.test(authorsRaw)) {
    problems.push({
      code: "AMPERSAND_MISSING",
      message: "Multiple authors should be joined with an ampersand (&) before the final author.",
    });
  }

  // --- Title ----------------------------------------------------------
  let title: string | null = null;
  if (yearMatch) {
    const afterYear = raw.slice(yearMatch.index + yearMatch[0].length).replace(/^\.\s*/, "");
    // Title = up to first sentence-ending period not part of an abbreviation.
    const t = /^((?:[^.?!]|\.(?=[A-Za-z]\.)|(?<=\b(?:vs|St|Mr|Dr|U\.S|e\.g|i\.e))\.)+[.?!]?)/.exec(afterYear);
    if (t) title = t[1]!.replace(/[.?!]\s*$/, "").trim() || null;
  }

  // --- DOI / URL ------------------------------------------------------
  const doiMatch = DOI_RE.exec(raw);
  let doi: string | null = null;
  if (doiMatch) {
    doi = doiMatch[1]!.replace(/[.,;]+$/, "");
    const matched = doiMatch[0];
    if (/^doi:\s*/i.test(matched) || /^DOI:\s*/i.test(matched)) {
      problems.push({
        code: "DOI_OLD_FORMAT",
        message: "DOIs should be presented as https://doi.org/xxxx URLs.",
        before: matched.replace(/[.,;]+$/, ""),
        after: `https://doi.org/${doi}`,
      });
    } else if (/^https?:\/\/dx\.doi\.org\//i.test(matched)) {
      problems.push({
        code: "DOI_OLD_FORMAT",
        message: "Use the modern https://doi.org/ form rather than dx.doi.org.",
        before: matched.replace(/[.,;]+$/, ""),
        after: `https://doi.org/${doi}`,
      });
    }
  }
  const urlMatch = doi ? null : URL_RE.exec(raw);
  const url = urlMatch ? urlMatch[0].replace(/[.,;]+$/, "") : null;

  if (/Retrieved from https?:\/\//i.test(raw) && !/Retrieved \w+ \d/.test(raw)) {
    problems.push({
      code: "URL_RETRIEVED_FROM",
      message:
        `APA 7 omits "Retrieved from" before URLs unless a retrieval date is needed.`,
      before: "Retrieved from ",
      after: "",
    });
  }

  // --- Journal info ---------------------------------------------------
  let journal: ParsedReference["journal"];
  // Pattern: Journal Name, 12(3), 45–67.
  const jm = /([A-Z][^,.]{2,80}),\s*(\d{1,4})\s*(?:\((\d{1,4}[A-Za-z]?)\))?\s*,\s*([\divxlc]+(?:\s*[–\-—]\s*[\divxlc]+)?)/.exec(
    raw
  );
  if (jm) {
    journal = {
      name: jm[1]!.trim(),
      volume: jm[2] ?? null,
      issue: jm[3] ?? null,
      pages: jm[4] ?? null,
    };
  }

  const { type, confidence } = classifyReference(raw, { doi, url, journal, isOrg });

  return {
    paragraphIndex,
    raw,
    authorsRaw,
    surnames,
    isOrganizationAuthor: isOrg,
    year,
    title,
    doi,
    url,
    type,
    typeConfidence: confidence,
    journal,
    problems,
  };
}

function extractSurnames(authorsRaw: string): { surnames: string[]; isOrg: boolean } {
  if (!authorsRaw) return { surnames: [], isOrg: false };
  // "Surname, I. I." pattern → personal authors
  const personal = /^[A-ZÀ-Þ][\w'’\-À-ÿ]+,\s*(?:[A-Z]\.\s*)+/.test(authorsRaw);
  if (!personal) {
    // Organization author, e.g. "American Psychological Association"
    return { surnames: [authorsRaw.replace(/\.$/, "").toLowerCase()], isOrg: true };
  }
  const surnames: string[] = [];
  // Split on separators between author units. Each unit: "Surname, I. I."
  const re = /([A-ZÀ-Þ][\w'’\-À-ÿ]+(?:\s[A-ZÀ-Þ][\w'’\-À-ÿ]+)?),\s*(?:[A-Z]\.?[\s-]*)+/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(authorsRaw)) !== null) {
    surnames.push(m[1]!.toLowerCase());
  }
  if (surnames.length === 0) surnames.push(authorsRaw.toLowerCase());
  return { surnames, isOrg: false };
}

function classifyReference(
  raw: string,
  info: {
    doi: string | null;
    url: string | null;
    journal?: ParsedReference["journal"];
    isOrg: boolean;
  }
): { type: ReferenceType; confidence: "high" | "medium" | "low" } {
  const r = raw;
  if (/\[Video\]|\bYouTube\b/i.test(r)) return { type: "video", confidence: "high" };
  if (/\[Audio podcast/i.test(r) || /\bpodcast\b/i.test(r))
    return { type: "podcast", confidence: /\[Audio podcast/i.test(r) ? "high" : "medium" };
  if (/\[Data set\]|\[Dataset\]/i.test(r)) return { type: "dataset", confidence: "high" };
  if (/\[Computer software\]|\[Software\]|\[Mobile app/i.test(r))
    return { type: "software", confidence: "high" };
  if (/\[(Doctoral dissertation|Doctoral thesis)/i.test(r))
    return { type: "dissertation", confidence: "high" };
  if (/\[(Master'?s thesis|Bachelor'?s thesis)/i.test(r))
    return { type: "thesis", confidence: "high" };
  if (/\[Paper presentation\]|\[Conference (session|presentation)\]|Proceedings of/i.test(r))
    return { type: "conference_paper", confidence: "high" };
  if (/\(Report No\.|Technical Report|\[Report\]/i.test(r))
    return { type: "report", confidence: "high" };
  if (/\bIn\s+[A-Z].*\(Eds?\.\),/.test(r)) return { type: "book_chapter", confidence: "high" };
  if (/\(Eds?\.\)\./.test(r)) return { type: "edited_book", confidence: "high" };
  if (info.journal && (info.doi || info.journal.volume)) {
    return { type: "journal_article", confidence: info.doi ? "high" : "medium" };
  }
  if (/(The New York Times|The Washington Post|The Guardian|Wall Street Journal|Times|Post|Herald|Tribune)\b/.test(r) && info.url) {
    return { type: "newspaper_article", confidence: "medium" };
  }
  if (/Government Printing Office|Government Publishing|U\.S\. Department|Department of|Bureau of|Ministry of/i.test(r) && info.isOrg) {
    return { type: "government_publication", confidence: "medium" };
  }
  if (info.isOrg && info.url) return { type: "organization_publication", confidence: "medium" };
  if (info.url && !info.doi) return { type: "webpage", confidence: "medium" };
  if (info.doi) return { type: "journal_article", confidence: "medium" };
  // Ends with "Publisher." pattern → book
  if (/\)\.\s+[^.]+\.\s+[A-Z][\w&' ]+(?: Press| Books| Publishing| Publishers| Wiley| Springer| Routledge| Sage| Pearson| Guilford| Norton| Penguin| Elsevier| Academic)\.?$/.test(r)) {
    return { type: "book", confidence: "medium" };
  }
  if (/\(\d+(?:st|nd|rd|th) ed\.\)/.test(r)) return { type: "book", confidence: "medium" };
  return { type: "unknown", confidence: "low" };
}
