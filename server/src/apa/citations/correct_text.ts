import { parseCitations } from "./parser.js";
import { deterministicFix } from "../rules/citations.js";

export interface TextCitationResult {
  text: string;
  citationsFound: number;
  changes: { before: string; after: string; reason: string }[];
  warnings: string[];
}

const surname = /^(?:(?:van|von|de|del|der|di|la|le|al|bin|ter)\s+)?[A-ZÀ-Þ][\p{L}'’–-]+$/iu;
const personalNames = (names: string[]) => names.length > 0 && names.every((name) => surname.test(name));

/** Plain-text mechanics only. Never adds sources, years, or citation locations. */
export function correctCitationText(input: string): TextCitationResult {
  const changes: TextCitationResult["changes"] = [];
  const warnings = new Set<string>();
  const record = (before: string, after: string, reason: string) => {
    if (before !== after) changes.push({ before, after, reason });
    return after;
  };
  // Preserve a pasted reference list; this tool works on in-text citations.
  const referenceStart = /^(?:references|bibliography|works cited)\s*:?\s*$/im.exec(input)?.index ?? input.length;
  let text = input.slice(0, referenceStart);
  const references = input.slice(referenceStart);
  const authorGroups = new Map<string, Set<string>>();
  for (const citation of parseCitations(text, 0)) {
    if (citation.authors.length < 3 || !personalNames(citation.authors)) continue;
    const key = `${citation.authors[0]}|${citation.year}`;
    const groups = authorGroups.get(key) ?? new Set<string>();
    groups.add(citation.authors.join("|"));
    authorGroups.set(key, groups);
  }
  const correct = (raw: string, kind: "parenthetical" | "narrative") => {
    let value = raw;
    const citation = parseCitations(raw, 0).find((c) => c.kind === kind);
    if (!citation) return raw;
    for (const problem of citation.problems) {
      if ((problem.code === "AND_IN_PARENTHETICAL" || problem.code === "AMPERSAND_IN_NARRATIVE") &&
        !personalNames(citation.authors)) {
        warnings.add(`Check the group-author name in ${raw}; its wording was preserved.`);
        continue;
      }
      const next = deterministicFix(value, problem.code);
      if (next) value = next;
      else warnings.add(problem.message);
    }
    // Three personal authors can be abbreviated mechanically. Ambiguous
    // author lists must retain enough names to distinguish their sources.
    if (!citation.etAl && citation.authors.length >= 3 && personalNames(citation.authors)) {
      if ((authorGroups.get(`${citation.authors[0]}|${citation.year}`)?.size ?? 0) > 1) {
        warnings.add(`Check ${raw}: multiple author lists would shorten to the same author and year, so the names were retained for disambiguation.`);
      } else if (kind === "parenthetical") {
        const yearIndex = value.indexOf(citation.year);
        value = `(${citation.authors[0]} et al., ${value.slice(yearIndex)}`;
      } else {
        value = `${citation.authors[0]} et al. ${value.slice(value.indexOf("("))}`;
      }
    }
    return record(raw, value, "Correct APA author–year punctuation and citation mechanics.");
  };
  text = text.replace(/\([^()\r\n]+\)/g, (group: string) => {
    if (!/(?:1[6-9]|20)\d{2}|n\.d\.|in press/i.test(group)) return group;
    // Formatting a year-only narrative group does not need author edits here.
    const normalized = group.replace(/,\s*(?=(?:1[6-9]|20)\d{2}[a-z]?\b|n\.d\.)/gi, ", ");
    const segments = normalized.slice(1, -1).split(";");
    const fixed = segments.map((segment) => {
      const leading = /^\s*/.exec(segment)![0];
      const trailing = /\s*$/.exec(segment)![0];
      return leading + correct(`(${segment.trim()})`, "parenthetical").slice(1, -1) + trailing;
    }).join(";");
    const result = `(${fixed})`;
    // Segment edits already have records; also record spacing when applicable.
    if (normalized !== group) record(group, normalized, "Add a space after the author–year comma.");
    return result;
  });
  // Process narrative spans from the end to preserve repeated occurrences.
  const narrative = parseCitations(text, 0).filter((c) => c.kind === "narrative");
  const seen = new Set<string>();
  for (const citation of narrative) {
    if (seen.has(citation.raw)) continue;
    seen.add(citation.raw);
    const fixed = correct(citation.raw, "narrative");
    if (fixed !== citation.raw) text = text.split(citation.raw).join(fixed);
  }
  // Accept a standalone citation typed without parentheses, but do not wrap
  // author/year text embedded in prose or full reference-list entries.
  text = text.replace(/^[^\r\n]+$/gm, (line) => {
    const trimmed = line.trim();
    const normalized = trimmed.replace(/,(?=(?:1[6-9]|20)\d{2})/, ", ");
    if (trimmed.includes("(") || /^(?:In|During|Since|Until|Year|From)\b/i.test(trimmed) ||
      !/^[\p{L}'’.,&\s–-]+,?\s+(?:(?:1[6-9]|20)\d{2}[a-z]?|n\.d\.)$/u.test(normalized)) return line;
    const citation = parseCitations(`(${normalized})`, 0)[0];
    if (!citation || !personalNames(citation.authors)) return line;
    const fixed = correct(`(${normalized})`, "parenthetical");
    return record(line, line.replace(trimmed, fixed), "Enclose a standalone author–year citation in parentheses.");
  });
  const citationsFound = parseCitations(text, 0).length;
  if (!citationsFound) warnings.add("No recognizable author–year citations were found. Include the author and year, for example (Smith, 2024).");
  if (references) warnings.add("The reference list was kept unchanged. This tool corrects in-text citations.");
  if ((text.match(/\(/g)?.length ?? 0) !== (text.match(/\)/g)?.length ?? 0)) warnings.add("Check unmatched parentheses; their intended positions could not be determined.");
  return { text: text + references, citationsFound, changes, warnings: [...warnings] };
}
