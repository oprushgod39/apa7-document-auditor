import { createHash, randomUUID } from "node:crypto";
import type { DocumentModel } from "../../docx/model.js";
import { childW, descendantsW, getAttrW, NS, setAttrW } from "../../docx/xml.js";
import { analyzeDocument } from "../analysis.js";
import type { ApaRule, Change } from "../types.js";

const PREFIX = "SWDraft_";
const RULE = "APA-CITATION-DRAFT";
const words = (s: string) => new Set(s.toLowerCase().match(/[\p{L}]{4,}/gu) ?? []);

export function clearConfirmedDrafts(model: DocumentModel, resolutions: Map<string, { optionId: string }>): void {
  for (const p of model.paragraphs) {
    for (const start of descendantsW(p.el, "bookmarkStart")) {
      const name = getAttrW(start, "name");
      if (!name?.startsWith(PREFIX) || ![...resolutions].some(([key, value]) =>
        key.endsWith(`|${name}`) && value.optionId === "confirmed")) continue;
      const id = getAttrW(start, "id");
      let node = start.nextSibling;
      while (node) {
        const next = node.nextSibling;
        if (node.nodeType === 1) {
          const el = node as typeof start;
          if (el.localName === "bookmarkEnd" && getAttrW(el, "id") === id) {
            el.parentNode!.removeChild(el);
            break;
          }
          for (const h of descendantsW(el, "highlight")) h.parentNode!.removeChild(h);
        }
        node = next;
      }
      start.parentNode!.removeChild(start);
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
  const bodies = model.paragraphs.filter((p) =>
    p.index >= analysis.bodyStartIndex &&
    (analysis.referencesHeadingIndex == null || p.index < analysis.referencesHeadingIndex) &&
    !p.insideTable && !p.hasDrawing && p.text.trim().split(/\s+/).length >= 12 &&
    !analysis.headings.some((h) => h.paragraphIndex === p.index));
  if (!bodies.length) return [];
  const counts = new Map<number, number>();
  let bookmarkId = Math.max(0, ...descendantsW(model.body, "bookmarkStart")
    .map((b) => Number(getAttrW(b, "id")) || 0)) + 1;
  const changes: Change[] = [];
  for (const usage of analysis.matching.referenceUsage) {
    if (usage.cited) continue;
    const ref = analysis.references[usage.referenceIndex]!;
    if (!ref.year || !ref.surnames.length) continue;
    const name = PREFIX + createHash("sha256").update(ref.raw).digest("hex").slice(0, 24);
    if (descendantsW(model.body, "bookmarkStart").some((b) => getAttrW(b, "name") === name)) continue;
    const confirmed = [...resolutions].some(([key, value]) => key.endsWith(`|${name}`) && value.optionId === "confirmed");
    const sourceWords = words(`${ref.authorsRaw} ${ref.title ?? ""}`);
    const ranked = bodies.map((p) => ({ p,
      score: [...words(p.text)].filter((w) => sourceWords.has(w)).length,
      used: counts.get(p.index) ?? 0,
    })).sort((a, b) => b.score - a.score || a.used - b.used || a.p.index - b.p.index);
    const target = ranked[0]!.p;
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
    // A yellow highlight and named bookmark survive downloading/reuploading.
    if (!confirmed) {
      const highlight = doc.createElementNS(NS.w, "w:highlight");
      setAttrW(highlight, "val", "yellow");
      properties.appendChild(highlight);
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
    counts.set(target.index, (counts.get(target.index) ?? 0) + 1);
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
          explanation: `This paragraph contains a yellow-highlighted draft citation: “${p.text.slice(0, 160)}”. Confirm only after checking that the source supports the passage. Regenerate after confirming to remove the highlight.`,
          location: { paragraphIndex: p.index }, originalValue: name, confidence: 0,
          autoFixable: false, userResolutionRequired: true,
          resolutionOptions: [{ id: "confirmed", label: "I checked and confirm this placement" }] });
      }
    }
    return { ruleId: RULE, checked: count, passed: 0, fixed: false,
      status: count ? "user_review" : "not_applicable" };
  },
};
