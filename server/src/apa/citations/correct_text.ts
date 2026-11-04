import { parseCitations } from "./parser.js";
import { deterministicFix } from "../rules/citations.js";
import { parseReference, type ParsedReference } from "../references/parser.js";

export interface TextCitationResult {
  text: string;
  citationsFound: number;
  changes: { before: string; after: string; reason: string }[];
  warnings: string[];
  uncitedParagraphs: { paragraph: number; excerpt: string }[];
  referencesFound: number;
}

const surname = /^(?:(?:van|von|de|del|der|di|la|le|al|bin|ter)\s+)?[A-ZÀ-Þ][\p{L}'’–-]+$/iu;
const personalNames = (names: string[]) => names.length > 0 && names.every((name) => surname.test(name));

/** Plain-text mechanics only. Never adds sources, years, or citation locations. */
export function correctCitationText(input: string, suppliedReferences?: ParsedReference[], options: { wrapStandalone?: boolean } = {}): TextCitationResult {
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
  const referenceEntries = suppliedReferences ?? references.split(/\r?\n/).slice(1)
    .filter((line) => line.trim()).map((line, index) => parseReference(line.trim(), index));
  // Recognize year:page notation before using the APA parser. Validate the
  // normalized group with author context so bare times/ratios stay intact.
  text = text.replace(/\([^()\r\n]+\)/g, (group: string, offset: number) => {
    const normalized = group.replace(/\b((?:1[6-9]|20)\d{2}[a-z]?)\s*:\s*(\d+(?:\s*[–—-]\s*\d+)?)(?=\s*[;)])/g,
      (_match, year: string, pages: string) => {
        const range = /[–—-]/.test(pages);
        return `${year}, ${range ? "pp." : "p."} ${pages.replace(/\s*[–—-]\s*/g, "–")}`;
      });
    if (normalized === group) return group;
    const context = text.slice(Math.max(0, offset - 120), offset) + normalized;
    const recognized = parseCitations(context, 0).some((c) =>
      c.groupRaw === normalized || (c.kind === "narrative" && c.raw.endsWith(normalized)));
    return recognized ? record(group, normalized, "Convert year:page notation to the APA year, p./pp. format.") : group;
  });
  // An abbreviation definition is not an author–year citation. Surface likely
  // group-author attributions without filling in an unknown publication date.
  const groupAuthor = /([A-Z][\p{L}’'.-]*(?:\s+(?:[A-Z][\p{L}’'.-]*|of|for|and|the|in|on|&)){1,12})\s*\(([A-Z][A-Z0-9&]{1,11})\)/gu;
  for (const match of [...text.matchAll(groupAuthor)].reverse()) {
    if (!/\b(?:Institute|University|Agency|Department|Association|Organization|Organisation|Council|Bureau|Commission|Administration|Center|Centre|Fund|Foundation)\b/i.test(match[1]!)) continue;
    const following = text.slice(match.index! + match[0].length);
    if (/^\s*\(\s*(?:(?:1[6-9]|20)\d{2}|n\.d\.|in press)/i.test(following)) continue;
    const normalized = (value: string) => value.toLowerCase().replace(/&/g, "and").replace(/[^\p{L}\p{N}]/gu, "");
    const fullName = normalized(match[1]!);
    const candidates = referenceEntries.filter((ref) => ref.isOrganizationAuthor && ref.year &&
      normalized(ref.authorsRaw).length > 10 && fullName.endsWith(normalized(ref.authorsRaw)));
    if (candidates.length === 1) {
      const updated = match[0].replace(`(${match[2]})`, `(${match[2]}, ${candidates[0]!.year})`);
      text = text.slice(0, match.index!) + record(match[0], updated, "Complete a group-author citation using its uniquely matching reference entry.") + text.slice(match.index! + match[0].length);
      continue;
    }
    warnings.add(`Possible missing citation year: ${match[0]}. If this is a source attribution, supply its publication year; the abbreviation alone is not a dated citation.`);
  }
  const authorGroups = new Map<string, Set<string>>();
  for (const ref of referenceEntries) {
    if (!ref.year || ref.surnames.length < 3 || ref.isOrganizationAuthor) continue;
    const key = `${ref.surnames[0]!.toLowerCase()}|${ref.year}`;
    const groups = authorGroups.get(key) ?? new Set<string>();
    groups.add(ref.surnames.join("|").toLowerCase());
    authorGroups.set(key, groups);
  }
  for (const citation of parseCitations(text, 0)) {
    if (citation.authors.length < 3 || !personalNames(citation.authors)) continue;
    const key = `${citation.authors[0]!.toLowerCase()}|${citation.year}`;
    const groups = authorGroups.get(key) ?? new Set<string>();
    groups.add(citation.authors.join("|").toLowerCase());
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
      if ((authorGroups.get(`${citation.authors[0]!.toLowerCase()}|${citation.year}`)?.size ?? 0) > 1) {
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
  if (options.wrapStandalone !== false) text = text.replace(/^[^\r\n]+$/gm, (line) => {
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
  const uncitedParagraphs = text.split(/\n\s*\n|\n/).map((paragraph, index) => ({ paragraph: index + 1, value: paragraph.trim() }))
    .filter(({ value }) => value.split(/\s+/).length >= 12 && parseCitations(value, 0).length === 0)
    .map(({ paragraph, value }) => ({ paragraph, excerpt: value.slice(0, 180) }));
  if (!citationsFound) warnings.add("No recognizable author–year citations were found. Include the author and year, for example (Smith, 2024).");
  if (references) warnings.add("The reference list was kept unchanged. This tool corrects in-text citations.");
  if ((text.match(/\(/g)?.length ?? 0) !== (text.match(/\)/g)?.length ?? 0)) warnings.add("Check unmatched parentheses; their intended positions could not be determined.");
  return { text: text + references, citationsFound, changes, warnings: [...warnings], uncitedParagraphs, referencesFound: referenceEntries.length };
}
