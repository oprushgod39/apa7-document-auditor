import type { ParsedReference } from "../../references/parser.js";
import { parseCitations } from "../parser.js";
import type { ClaimDecision, ClaimReport, EvidenceSource } from "./types.js";

export function citationLabel(ref: ParsedReference): string | null {
  if (!ref.year || !ref.surnames.length) return null;
  const authors = ref.isOrganizationAuthor ? ref.authorsRaw.replace(/\.$/, "").trim() : ref.surnames.map((s) => {
    const at = ref.authorsRaw.toLowerCase().indexOf(s.toLowerCase());
    return at < 0 ? s : ref.authorsRaw.slice(at, at + s.length);
  });
  return `${typeof authors === "string" ? authors : authors.length > 2 ? `${authors[0]} et al.` : authors.join(" & ")}, ${ref.year}`;
}
const normalize = (s: string) => s.replace(/\s+/g, " ").trim();
export function planEdits(text: string, paragraph: number, decisions: ClaimDecision[], sources: EvidenceSource[], editable = true) {
  const edits: { paragraph: number; start: number; length: number; value: string }[] = [];
  const claims: ClaimReport[] = [];
  for (const d of decisions) {
    const report: ClaimReport = { ...d, paragraph, correctedCitation: "", applied: false, review: "" };
    claims.push(report);
    const reject = (reason: string) => { report.review = reason; };
    if (!text.includes(d.sentence) || !d.sentence.includes(d.claim)) { report.status = "insufficient"; reject("Model claim did not map to the original text."); continue; }
    const verified = d.sourceIds.length > 0 && d.sourceIds.every((id) => {
      const source = sources.find((s) => s.id === id);
      const content = source?.text.includes("Abstract:") ? source.text.split("Abstract:").slice(1).join("Abstract:") : source?.text ?? "";
      return source?.status === "available" && d.evidence.some((e) => e.sourceId === id && e.identityConfirmed && e.quote.length >= 25 &&
        !normalize(source.reference.title ?? "").includes(normalize(e.quote)) && normalize(content).includes(normalize(e.quote)));
    });
    if (d.status === "supported" && !verified) { report.status = "insufficient"; reject("No valid verbatim evidence from an accessible, identified source."); continue; }
    if (d.action === "none") continue;
    if (d.status !== "supported" || !verified) { reject("Evidence is insufficient for automatic correction."); continue; }
    if (d.category === "statistic") {
      const numbers = d.claim.match(/\d+(?:[.,]\d+)*(?:%|\b)/g) ?? [];
      const quotations = d.evidence.filter((e) => d.sourceIds.includes(e.sourceId)).map((e) => e.quote).join(" ");
      if (numbers.some((n) => !quotations.includes(n))) { report.status = "insufficient"; reject("The cited evidence does not contain the claim's numerical values."); continue; }
    }
    if (!editable) { reject("Paragraph contains protected fields or revisions; manual correction required."); continue; }
    if (d.category === "quotation" || /\bas cited in\b/i.test(d.sentence)) { reject("Quotation locator or secondary citation needs manual verification."); continue; }
    const selected = d.sourceIds.map((id) => sources.find((s) => s.id === id)!);
    const labels = selected.map((s) => citationLabel(s.reference));
    if (labels.some((label) => !label || sources.filter((s) => citationLabel(s.reference) === label).length !== 1)) { reject("Missing or ambiguous author–year metadata; no citation was assigned."); continue; }
    const citation = `(${[...new Set(labels as string[])].sort((a, b) => a.localeCompare(b)).join("; ")})`;
    let start: number, length = 0, value = citation;
    if (d.action === "insert") {
      if (!d.anchor || !d.sentence.includes(d.anchor) || text.indexOf(d.anchor) !== text.lastIndexOf(d.anchor)) { reject("Insertion location is ambiguous."); continue; }
      start = text.indexOf(d.anchor) + d.anchor.length;
      if (/[.!?]$/.test(d.anchor) || !/^[,;:.!?\s)]|^$/.test(text.slice(start))) { reject("Citation anchor is not at a clause boundary."); continue; }
      if (d.sentence.includes(citation)) { reject("The proposed citation is already present."); continue; }
      value = ` ${citation}`;
    } else {
      const original = d.originalCitation;
      if (!original || !d.sentence.includes(original) || text.indexOf(original) !== text.lastIndexOf(original) ||
        !parseCitations(text, paragraph).some((c) => c.kind === "parenthetical" && c.groupRaw === original)) { reject("Existing citation cannot be safely isolated."); continue; }
      start = text.indexOf(original); length = original.length;
      // Removing an attribution needs whole-paragraph scope validation; retain it.
      if (d.action === "remove") { reject("Redundant/misplaced citation removal requires attribution review."); continue; }
    }
    if (edits.some((e) => start <= e.start + e.length && start + length >= e.start)) { reject("Conflicting edits at this claim; review required."); continue; }
    edits.push({ paragraph, start, length, value }); report.correctedCitation = citation; report.applied = true;
  }
  return { claims, edits };
}
