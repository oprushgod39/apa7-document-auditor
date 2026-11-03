import { createHash, randomUUID } from "node:crypto";
import type { DocumentModel } from "../../docx/model.js";
import { childW, descendantsW, getAttrW, NS, setAttrW } from "../../docx/xml.js";
import { analyzeDocument } from "../analysis.js";
import type { ApaRule, Change } from "../types.js";

const PREFIX = "SWDraft_";
const RULE = "APA-CITATION-DRAFT";
const words = (s: string) => new Set(s.toLowerCase().match(/[\p{L}]{4,}/gu) ?? []);
const TARGET = 3;

function bodyParagraphs(model: DocumentModel, analysis: ReturnType<typeof analyzeDocument>) {
  return model.paragraphs.filter((p) =>
    p.index >= analysis.bodyStartIndex &&
    (analysis.referencesHeadingIndex == null || p.index < analysis.referencesHeadingIndex) &&
    !p.insideTable && !p.isEmpty &&
    (p.text.trim().split(/\s+/).length >= 12 || /[.!?]$/.test(p.text.trim())) &&
    p.text.trim() !== analysis.detectedMetadata.title?.trim() &&
    !/^(?:table|figure)\s+\d+\b|^note\.\s/i.test(p.text.trim()) &&
    !analysis.headings.some((h) => h.paragraphIndex === p.index));
}

export function clearConfirmedDrafts(model: DocumentModel, resolutions: Map<string, { optionId: string }>): void {
  for (const p of model.paragraphs) {
    for (const start of descendantsW(p.el, "bookmarkStart")) {
      const name = getAttrW(start, "name");
      if (!name?.startsWith(PREFIX)) continue;
      const confirmed = [...resolutions].some(([key, value]) =>
        key.endsWith(`|${name}`) && value.optionId === "confirmed");
      const id = getAttrW(start, "id");
      let node = start.nextSibling;
      while (node) {
        const next = node.nextSibling;
        if (node.nodeType === 1) {
          const el = node as typeof start;
          if (el.localName === "bookmarkEnd" && getAttrW(el, "id") === id) {
            if (confirmed) el.parentNode!.removeChild(el);
            break;
          }
          for (const h of descendantsW(el, "highlight")) h.parentNode!.removeChild(h);
        }
        node = next;
      }
      if (confirmed) start.parentNode!.removeChild(start);
      model.pkg.markDirty("word/document.xml");
    }
  }
}

/** Draft placements are proposals, never evidence that a source supports a claim. */
export function insertDraftCitations(
  model: DocumentModel,
  resolutions: Map<string, { optionId: string }>,
): Change[] {
  const analysis = analyzeDocument(model);
  if (analysis.annotatedBibliography) return [];
  const bodies = bodyParagraphs(model, analysis);
  if (!bodies.length) return [];
  let bookmarkId = Math.max(0, ...descendantsW(model.body, "bookmarkStart")
    .map((b) => Number(getAttrW(b, "id")) || 0)) + 1;
  const changes: Change[] = [];
  const placements = bodies.flatMap((target) => {
    const citations = analysis.citations.filter((c) => c.paragraphIndex === target.index);
    const existing = new Set(analysis.matching.citationMatches
      .filter((match) => analysis.citations[match.citationIndex]?.paragraphIndex === target.index)
      .flatMap((match) => match.referenceIndexes));
    const targetWords = words(target.text);
    return analysis.references.map((ref, index) => ({ ref, index,
      score: [...words(`${ref.authorsRaw} ${ref.title ?? ""}`)].filter((w) => targetWords.has(w)).length,
    })).filter(({ ref, index }) => ref.year && ref.surnames.length && !existing.has(index))
      .sort((a, b) => b.score - a.score || a.index - b.index)
      .slice(0, Math.max(0, TARGET - citations.length))
      .map(({ ref }) => ({ ref, target }));
  });
  for (const { ref, target } of placements) {
    const name = PREFIX + createHash("sha256").update(`${ref.raw}|${target.index}`).digest("hex").slice(0, 24);
    if (descendantsW(model.body, "bookmarkStart").some((b) => getAttrW(b, "name") === name)) continue;
    const confirmed = [...resolutions].some(([key, value]) => key.endsWith(`|${name}`) && value.optionId === "confirmed");
    const author = ref.isOrganizationAuthor ? ref.authorsRaw.trim().replace(/\.$/, "")
      : ref.surnames.map((s) => {
        // Recover the author's original spelling/capitalization from the reference.
        const start = ref.authorsRaw.toLowerCase().indexOf(s.toLowerCase());
        return start >= 0 ? ref.authorsRaw.slice(start, start + s.length) : s;
      });
    const authors = typeof author === "string" ? author
      : author.length >= 3 ? `${author[0]} et al.` : author.join(" & ");
    const citation = `(${authors}, ${ref.year})`;
    const before = descendantsW(target.el, "t").map((t) => t.textContent ?? "").join("");
    // Keep the existing runs and formatting; move only the terminal full stop.
    const texts = descendantsW(target.el, "t");
    const last = [...texts].reverse().find((t) => (t.textContent ?? "").trim().length > 0);
    let punctuation = "";
    if (last && /[.!?]\s*$/.test(last.textContent ?? "")) {
      punctuation = /([.!?])\s*$/.exec(last.textContent!)![1]!;
      last.textContent = last.textContent!.replace(/[.!?]\s*$/, "");
    }
    const doc = model.documentXml;
    const run = doc.createElementNS(NS.w, "w:r");
    const lastRun = descendantsW(target.el, "r").at(-1);
    const properties = childW(lastRun, "rPr")?.cloneNode(true) ?? doc.createElementNS(NS.w, "w:rPr");
    run.appendChild(properties);
    for (const highlight of descendantsW(run, "highlight")) highlight.parentNode!.removeChild(highlight);
    // An invisible bookmark tracks review status without coloring document text.
    if (!confirmed) {
      const start = doc.createElementNS(NS.w, "w:bookmarkStart");
      setAttrW(start, "id", String(bookmarkId));
      setAttrW(start, "name", name);
      target.el.appendChild(start);
    }
    const text = doc.createElementNS(NS.w, "w:t");
    text.setAttribute("xml:space", "preserve");
    text.textContent = ` ${citation}${punctuation}`;
    run.appendChild(text);
    target.el.appendChild(run);
    if (!confirmed) {
      const end = doc.createElementNS(NS.w, "w:bookmarkEnd");
      setAttrW(end, "id", String(bookmarkId++));
      target.el.appendChild(end);
    }
    model.pkg.markDirty("word/document.xml");
    changes.push({ id: randomUUID(), ruleId: RULE, category: "citations",
      location: { paragraphIndex: target.index }, before,
      after: descendantsW(target.el, "t").map((t) => t.textContent ?? "").join(""),
      reason: confirmed ? "Draft citation placement confirmed by you."
        : "Unverified draft placement for manual review; source support has not been checked.",
      confidence: confirmed ? 1 : 0, stage: "format", timestamp: new Date().toISOString() });
  }
  return changes;
}

export const draftCitationRule: ApaRule = {
  id: RULE, category: "citations", description: "Draft citation placements require confirmation.",
  severity: "warning", applies: () => true,
  run(ctx) {
    let count = 0;
    for (const p of ctx.model.paragraphs) {
      for (const bookmark of descendantsW(p.el, "bookmarkStart")) {
        const name = getAttrW(bookmark, "name");
        if (!name?.startsWith(PREFIX)) continue;
        count++;
        const citationRun = bookmark.nextSibling;
        const citation = citationRun?.nodeType === 1
          ? descendantsW(citationRun as typeof bookmark, "t").map((t) => t.textContent ?? "").join("").trim()
          : "";
        ctx.addIssue({ ruleId: RULE, category: "citations", severity: "warning", status: "user_review",
          message: `Unverified draft citation ${citation} — check its source and location.`,
          explanation: `Proposed citation: ${citation}. Paragraph: “${p.text.slice(0, 160)}”. Confirm only after checking that the source supports the passage. Draft citations use normal text with no color highlighting.`,
          location: { paragraphIndex: p.index }, originalValue: name, confidence: 0,
          autoFixable: false, userResolutionRequired: true,
          resolutionOptions: [{ id: "confirmed", label: "I checked and confirm this placement" }] });
      }
    }
    if (ctx.settings.draftCitationPlacement && ctx.settings.mode !== "check" && !ctx.analysis.annotatedBibliography) {
      for (const p of bodyParagraphs(ctx.model, ctx.analysis)) {
        const found = ctx.analysis.citations.filter((c) => c.paragraphIndex === p.index).length;
        if (found >= TARGET) continue;
        count++;
        ctx.addIssue({ ruleId: RULE, category: "citations", severity: "warning", status: "user_review",
          message: `Paragraph has ${found} of the requested ${TARGET} citations.`,
          explanation: "There are not enough distinct references with usable authors and dates for this paragraph. Add suitable sources or review the requested citation count.",
          location: { paragraphIndex: p.index }, originalValue: `shortfall:${p.index}`,
          confidence: 1, autoFixable: false, userResolutionRequired: true,
          resolutionOptions: [{ id: "reviewed", label: "I reviewed the citation count" }] });
      }
    }
    return { ruleId: RULE, checked: count, passed: 0, fixed: false,
      status: count ? "user_review" : "not_applicable" };
  },
};
