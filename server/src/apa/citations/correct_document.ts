import { randomUUID } from "node:crypto";
import { diffChars } from "diff";
import { DocxPackage, verifyDocxIntegrity } from "../../docx/package.js";
import { buildDocumentModel, contentFingerprint } from "../../docx/model.js";
import { rawParagraphText, replaceParagraphTextAt } from "../../docx/text.js";
import { descendantsW } from "../../docx/xml.js";
import { assertContentPreserved } from "../../pipeline.js";
import { analyzeDocument } from "../analysis.js";
import type { Change } from "../types.js";
import { correctCitationText } from "./correct_text.js";
import { parseCitations } from "./parser.js";
import { matchCitationsToReferences } from "../references/matcher.js";

export async function correctCitationDocument(input: Buffer) {
  const pkg = await DocxPackage.load(input);
  const model = await buildDocumentModel(pkg);
  const analysis = analyzeDocument(model);
  const before = contentFingerprint(model);
  const changes: Change[] = [];
  const warnings = new Set<string>();
  const expected = new Map<number, string>();
  let citationsFound = 0;
  const used = new Set<number>();
  const aliases = new Map<string, Set<string>>();
  const documentText = model.paragraphs.map((p) => p.text).join("\n");
  for (const ref of analysis.references) {
    if (!ref.isOrganizationAuthor) continue;
    const author = ref.authorsRaw.trim().replace(/\.$/, "");
    const escaped = author.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const declaration = new RegExp(`${escaped}\\s*\\(([A-Z][A-Z0-9&]{1,11})(?:\\s*,\\s*(?:\\d{4}[a-z]?|n\\.d\\.))?\\)`, "g");
    for (const match of documentText.matchAll(declaration)) {
      const names = aliases.get(match[1]!) ?? new Set<string>();
      names.add(author); aliases.set(match[1]!, names);
    }
  }
  if (!analysis.references.length) warnings.add("No reference list was detected. Citation mechanics were checked without reference matching.");
  for (const p of model.paragraphs) {
    if (p.index < analysis.bodyStartIndex || p.isEmpty ||
      (analysis.referencesHeadingIndex != null && p.index >= analysis.referencesHeadingIndex) ||
      analysis.headings.some((h) => h.paragraphIndex === p.index)) continue;
    const original = rawParagraphText(p.el);
    if (original.length > 50_000 || descendantsW(p.el, "fldChar").length || descendantsW(p.el, "instrText").length || descendantsW(p.el, "ins").length || descendantsW(p.el, "del").length) {
      warnings.add(`Paragraph ${p.index + 1}: contains fields, tracked changes, or unusually long text; check its citations manually.`);
      continue;
    }
    // Disable standalone wrapping on body prose: an isolated author–year line
    // can be a heading rather than a citation.
    const result = correctCitationText(original, analysis.references, { wrapStandalone: false });
    for (const warning of result.warnings) {
      if (!warning.startsWith("No recognizable")) warnings.add(`Paragraph ${p.index + 1}: ${warning}`);
    }
    if (result.text !== original) {
      const diff = diffChars(original, result.text, { timeout: 150, maxEditLength: 10_000 });
      if (!diff) { warnings.add(`Paragraph ${p.index + 1}: corrections were too complex; its original text was retained.`); continue; }
      let offset = 0;
      const edits: { offset: number; length: number; value: string }[] = [];
      for (const part of diff) {
        if (!part.added && !part.removed) { offset += part.value.length; continue; }
        const previous = edits.at(-1);
        if (part.removed) { edits.push({ offset, length: part.value.length, value: "" }); offset += part.value.length; }
        else if (previous && previous.offset + previous.length === offset) previous.value += part.value;
        else edits.push({ offset, length: 0, value: part.value });
      }
      for (const edit of edits.reverse()) {
        if (!replaceParagraphTextAt(p.el, edit.offset, edit.length, edit.value)) throw new Error("Could not preserve a citation's text range.");
      }
      if (rawParagraphText(p.el) !== result.text) throw new Error("Citation text validation failed.");
      pkg.markDirty("word/document.xml");
      changes.push({ id: randomUUID(), ruleId: "APA-CITATION-TEXT", category: "citations",
        location: { paragraphIndex: p.index }, before: original, after: result.text,
        reason: "Correct citation mechanics using supplied references; preserve surrounding wording and document formatting.",
        confidence: 0.95, stage: "format", timestamp: new Date().toISOString() });
    }
    expected.set(p.index, result.text);
    const citations = parseCitations(result.text, p.index);
    citationsFound += citations.length;
    const expanded = citations.map((citation) => {
      const names = citation.authors.length === 1 ? aliases.get(citation.authors[0]!) : undefined;
      return names?.size === 1 ? { ...citation, authors: [[...names][0]!] } : citation;
    });
    const matches = matchCitationsToReferences(expanded, analysis.references);
    for (const match of matches.citationMatches) {
      if (match.status === "missing_reference" || match.status === "ambiguous") {
        warnings.add(`Paragraph ${p.index + 1}: ${match.detail}. Check the author and year against the reference list.`);
      }
      for (const index of match.referenceIndexes) used.add(index);
    }
  }
  const output = changes.length ? await pkg.save() : input;
  await verifyDocxIntegrity(output);
  const out = await buildDocumentModel(await DocxPackage.load(output));
  assertContentPreserved(before, contentFingerprint(out), changes);
  for (const p of out.paragraphs) {
    const anticipated = expected.get(p.index);
    if (anticipated != null && rawParagraphText(p.el) !== anticipated) throw new Error("Citation output did not match the expected text.");
  }
  const uncitedReferences = analysis.references.filter((_r, i) => !used.has(i)).map((r) => r.raw);
  return { output, summary: { citationsFound, referencesFound: analysis.references.length,
    paragraphsCorrected: changes.length, changes: changes.map((c) => ({ paragraph: c.location.paragraphIndex! + 1, before: c.before, after: c.after, reason: c.reason })),
    warnings: [...warnings], uncitedReferences } };
}
