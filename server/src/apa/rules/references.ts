import type { ApaRule } from "../types.js";
import {
  setParagraphAlignment,
  setParagraphIndent,
  setPageBreakBefore,
  setRunBold,
  setParagraphRunColorBlack,
  replaceParagraphRuns,
} from "../../docx/edit.js";
import { childrenW, paragraphText } from "../../docx/xml.js";
import { replaceParagraphText } from "../../docx/text.js";
import { referenceSortKey } from "../references/matcher.js";
import { result, loc, markDocDirty } from "./util.js";
import { excerptOf } from "../types.js";

/** H/I/J. References section rules. */

interface TextRange { start: number; end: number }

/** Identify journal title and volume ranges in a conventional journal reference. */
/**
 * Italic range for a standalone work (report, book, webpage) that is its
 * own container — no journal/volume to anchor on. APA 7 italicizes the
 * work's own title in this case (e.g. `Author. (Year). Title of the
 * report: Subtitle. Publisher. URL` — the title, not the publisher, is
 * italicized). The title is taken as running from right after the
 * year-parenthetical to the first sentence-ending period, which matches
 * standard APA reference punctuation (a colon separates title from
 * subtitle, not a period, so this doesn't cut a subtitle off early).
 *
 * Deliberately conservative: skipped entirely for anything that looks like
 * a book-chapter/edited-volume reference ("In Editor (Ed.), Book Title...")
 * where it's the *containing* book's title that should be italicized, not
 * the chapter title this heuristic would otherwise grab.
 */
function standaloneWorkItalicRange(text: string): TextRange | null {
  if (/\bIn\s+[A-Z][^()]*\(Eds?\.\)/.test(text)) return null;
  const year = /\((?:1[6-9]\d{2}|20\d{2})[a-z]?(?:,[^)]*)?\)\.\s+/i.exec(text);
  if (!year) return null;
  const titleStart = year.index + year[0].length;
  const period = /[.?!]\s+/.exec(text.slice(titleStart));
  if (!period) return null;
  const titleEnd = titleStart + period.index;
  if (titleEnd <= titleStart) return null;
  return { start: titleStart, end: titleEnd };
}

function journalItalicRanges(text: string): TextRange[] {
  const year = /\((?:1[6-9]\d{2}|20\d{2})[a-z]?\)\.\s+/i.exec(text);
  if (!year) return [];
  const tailStart = year.index + year[0].length;
  const volume = /,\s+(\d+)(?:\([^)]*\))?,\s+/.exec(text.slice(tailStart));
  if (!volume) return [];
  const comma = tailStart + volume.index;
  let titleEnd = -1;
  for (const match of text.slice(tailStart, comma).matchAll(/[.?!]\s+/g)) {
    titleEnd = tailStart + match.index!;
  }
  if (titleEnd < tailStart) return [];
  let journalStart = titleEnd + 2;
  while (/\s/.test(text[journalStart] ?? "")) journalStart++;
  let journalEnd = comma;
  while (journalEnd > journalStart && /\s/.test(text[journalEnd - 1] ?? "")) journalEnd--;
  const volumeDigits = volume[1]!;
  const volumeStart = comma + volume[0].indexOf(volumeDigits);
  return [
    { start: journalStart, end: journalEnd },
    { start: volumeStart, end: volumeStart + volumeDigits.length },
  ];
}

function rangesAreItalic(
  p: import("../../docx/model.js").ParagraphModel,
  ranges: TextRange[]
): boolean {
  let offset = 0;
  for (const run of p.runs) {
    const start = offset;
    const end = start + run.text.length;
    offset = end;
    if (ranges.some((range) => start < range.end && end > range.start) && run.effective.italic !== true) {
      return false;
    }
  }
  return true;
}

function segmentedReference(text: string, ranges: TextRange[]) {
  const boundaries = [...new Set([0, text.length, ...ranges.flatMap((r) => [r.start, r.end])])]
    .sort((a, b) => a - b);
  return boundaries.slice(0, -1).map((start, i) => {
    const end = boundaries[i + 1]!;
    return {
      text: text.slice(start, end),
      italic: ranges.some((range) => start >= range.start && end <= range.end),
      bold: false,
    };
  });
}

export const referenceRules: ApaRule[] = [
  {
    id: "APA-REFERENCE-001",
    category: "references",
    description: `The reference list is titled "References", bold, centered, on a new page.`,
    severity: "warning",
    applies: (ctx) =>
      ctx.analysis.referencesHeadingIndex != null ||
      ctx.analysis.embeddedReferencesHeadingCandidate != null,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      if (analysis.referencesHeadingIndex == null) {
        // In fix mode the pipeline splits this paragraph and re-analyzes
        // before rules run, so referencesHeadingIndex is only still null
        // here in check mode — report it instead of silently skipping.
        const cand = analysis.embeddedReferencesHeadingCandidate!;
        const p = model.paragraphs[cand.paragraphIndex]!;
        ctx.addIssue({
          ruleId: "APA-REFERENCE-001",
          category: "references",
          severity: "warning",
          status: "fail",
          message:
            `The "References" heading is combined with the preceding paragraph, separated ` +
            `only by a line break, instead of being on its own paragraph.`,
          explanation:
            `This usually happens when "References" was typed after pressing Shift+Enter ` +
            `instead of Enter. Formatting this document will split it onto its own paragraph.`,
          location: loc(p),
          originalValue: p.text.trim(),
          suggestedValue: "References",
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
        return result("APA-REFERENCE-001", 1, 0, false, "fail");
      }
      const p = model.paragraphs[analysis.referencesHeadingIndex!]!;
      const text = p.text.trim();
      const correctText = /^references$/i.test(text) && text === "References";
      const bold = p.runProps.bold === true;
      const centered = p.props.alignment === "center";
      const prev = model.paragraphs[p.index - 1];
      const newPage =
        p.props.pageBreakBefore === true ||
        (prev != null && prev.hasPageBreakAfterInRuns);
      if (correctText && bold && centered && newPage) {
        return result("APA-REFERENCE-001", 1, 1, false, null);
      }
      if (fix) {
        const beforeDesc = `"${text}", ${bold ? "bold" : "not bold"}, ${centered ? "centered" : "not centered"}, ${newPage ? "new page" : "not on new page"}`;
        if (!correctText) replaceParagraphText(p.el, text, "References");
        if (!centered) setParagraphAlignment(model.documentXml, p.el, "center");
        if (!bold) {
          for (const r of childrenW(p.el, "r")) setRunBold(model.documentXml, r, true);
        }
        setParagraphRunColorBlack(model.documentXml, p.el);
        if (!newPage) setPageBreakBefore(model.documentXml, p.el, true);
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-REFERENCE-001",
          category: "references",
          location: loc(p),
          before: beforeDesc,
          after: `"References", bold, centered, starting on a new page`,
          reason: `APA 7 titles the reference list "References" in bold, centered, on its own page.`,
          confidence: 0.95,
        });
        return result("APA-REFERENCE-001", 1, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-REFERENCE-001",
        category: "references",
        severity: "warning",
        status: "fail",
        message: `The reference list heading should be "References", bold, centered, on a new page.`,
        location: loc(p),
        originalValue: text,
        suggestedValue: "References",
        confidence: 0.95,
        autoFixable: true,
        userResolutionRequired: false,
      });
      return result("APA-REFERENCE-001", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-REFERENCE-008",
    category: "references",
    description: "Journal titles and volume numbers are italicized in journal references; standalone works (reports, books, webpages) have their own title italicized.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.referenceEntryIndexes.length > 0,
    run(ctx, fix) {
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const idx of ctx.analysis.referenceEntryIndexes) {
        const p = ctx.model.paragraphs[idx]!;
        const journalRanges = journalItalicRanges(p.text);
        const standaloneRange = journalRanges.length === 0 ? standaloneWorkItalicRange(p.text) : null;
        const ranges = journalRanges.length > 0 ? journalRanges : standaloneRange ? [standaloneRange] : [];
        if (ranges.length === 0) continue;
        checked++;
        if (rangesAreItalic(p, ranges)) {
          passed++;
          continue;
        }
        if (fix) {
          // Re-derive from the paragraph's live DOM text, not the cached
          // analysis snapshot (`p.text`): an earlier rule in this same fix
          // pass may have already edited this paragraph's runs, and
          // rebuilding from stale text would silently discard that edit.
          const liveText = paragraphText(p.el);
          const liveJournalRanges = journalItalicRanges(liveText);
          const liveStandalone = liveJournalRanges.length === 0 ? standaloneWorkItalicRange(liveText) : null;
          const liveRanges = liveJournalRanges.length > 0 ? liveJournalRanges : liveStandalone ? [liveStandalone] : [];
          const segments = segmentedReference(liveText, liveRanges).map((segment) => ({
            ...segment,
            font: ctx.req.font,
            halfPoints: ctx.req.fontSizePt * 2,
            black: true,
          }));
          if (liveRanges.length > 0 && replaceParagraphRuns(ctx.model.documentXml, p.el, segments)) {
            fixedCount++;
            markDocDirty(ctx);
            ctx.addChange({
              ruleId: "APA-REFERENCE-008",
              category: "references",
              location: loc(p),
              before: journalRanges.length > 0 ? "journal title and/or volume not italicized" : "work title not italicized",
              after: journalRanges.length > 0 ? "journal title and volume italicized" : "work title italicized",
              reason: journalRanges.length > 0
                ? "APA 7 italicizes the journal title and volume number in periodical references."
                : "APA 7 italicizes the title of a standalone work (report, book, webpage) that is its own container.",
              confidence: 0.9,
            });
            continue;
          }
        }
        anyFail = true;
        ctx.addIssue({
          ruleId: "APA-REFERENCE-008",
          category: "references",
          severity: "warning",
          status: "fail",
          message: "A journal reference is missing italics on its journal title or volume number.",
          location: loc(p),
          confidence: 0.85,
          autoFixable: false,
          userResolutionRequired: false,
        });
      }
      return result("APA-REFERENCE-008", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-REFERENCE-002",
    category: "references",
    description: "Reference entries use a 0.5-inch hanging indent.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.referenceEntryIndexes.length > 0,
    run(ctx, fix) {
      const { model, analysis, req } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const idx of analysis.referenceEntryIndexes) {
        const p = model.paragraphs[idx]!;
        checked++;
        const hanging = p.props.hangingIndent ?? 0;
        const firstLine = p.props.firstLineIndent ?? 0;
        const ok = Math.abs(hanging - req.hangingIndentTwips) <= 40 && firstLine === 0;
        if (ok) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphIndent(model.documentXml, p.el, {
            hanging: req.hangingIndentTwips,
            left: req.hangingIndentTwips,
            firstLine: null,
          });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-REFERENCE-002",
            category: "references",
            location: loc(p),
            before:
              hanging > 0
                ? `hanging indent ${(hanging / 1440).toFixed(2)}"`
                : firstLine > 0
                  ? `first-line indent ${(firstLine / 1440).toFixed(2)}"`
                  : "no hanging indent",
            after: '0.5" hanging indent',
            reason: "APA 7 reference entries use a 0.5-inch hanging indent.",
            confidence: 0.95,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-REFERENCE-002",
          category: "references",
          severity: "warning",
          status: "fail",
          message: `${checked - passed} reference entr(ies) lack a 0.5" hanging indent.`,
          confidence: 0.95,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-REFERENCE-002", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-REFERENCE-003",
    category: "references",
    description: "Reference entries are in alphabetical order.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.referenceEntryIndexes.length > 1,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const entries = analysis.references;
      const keys = entries.map(referenceSortKey);
      const sortedKeys = [...keys].sort((a, b) => a.localeCompare(b));
      const inOrder = keys.every((k, i) => k === sortedKeys[i]);
      if (inOrder) return result("APA-REFERENCE-003", entries.length, entries.length, false, null);

      // Only reorder automatically when every entry parsed with an author —
      // otherwise the sort key is unreliable.
      const allParsed = entries.every(
        (e) => e.surnames.length > 0 && e.authorsRaw.length > 0
      );
      if (fix && allParsed) {
        const order = entries
          .map((e, i) => ({ i, key: keys[i]! }))
          .sort((a, b) => a.key.localeCompare(b.key));
        const paraEls = analysis.referenceEntryIndexes.map(
          (idx) => model.paragraphs[idx]!.el
        );
        const anchor = paraEls[0]!;
        const parent = anchor.parentNode!;
        const marker = anchor.previousSibling;
        // Detach in current order, reinsert sorted at the same position.
        for (const el of paraEls) parent.removeChild(el);
        let insertAfter = marker;
        for (const { i } of order) {
          const el = paraEls[i]!;
          if (insertAfter && insertAfter.nextSibling) {
            parent.insertBefore(el, insertAfter.nextSibling);
          } else if (insertAfter) {
            parent.appendChild(el);
          } else {
            parent.insertBefore(el, parent.firstChild);
          }
          insertAfter = el;
        }
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-REFERENCE-003",
          category: "references",
          location: { description: "References section" },
          before: "Entries out of alphabetical order",
          after: "Entries sorted alphabetically by author surname and year",
          reason: "APA 7 reference lists are alphabetized by first author.",
          confidence: 0.9,
          documentWide: true,
        });
        return result("APA-REFERENCE-003", entries.length, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-REFERENCE-003",
        category: "references",
        severity: "warning",
        status: allParsed ? "fail" : "user_review",
        message: "Reference entries are not in alphabetical order.",
        explanation: allParsed
          ? undefined
          : "Some entries could not be parsed reliably, so automatic sorting was not applied. Please reorder manually.",
        confidence: allParsed ? 0.9 : 0.6,
        autoFixable: allParsed,
        userResolutionRequired: !allParsed,
        resolutionOptions: allParsed
          ? undefined
          : [
              { id: "acknowledge", label: "I will reorder manually", description: "Handle ordering in the document." },
              { id: "correct", label: "Order is correct", description: "The current order follows APA rules (e.g. same-author sequences)." },
            ],
      });
      return result(
        "APA-REFERENCE-003",
        entries.length,
        0,
        false,
        allParsed ? "fail" : "user_review"
      );
    },
  },

  {
    id: "APA-REFERENCE-004",
    category: "references",
    description: "DOIs and URLs use modern APA 7 presentation.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.references.length > 0,
    run(ctx, fix) {
      const { model, analysis, settings } = ctx;
      const doFix = fix && settings.fixCitationMechanics;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "warning" | null = null;
      for (const ref of analysis.references) {
        checked++;
        const fixableProblems = ref.problems.filter(
          (pr) =>
            (pr.code === "DOI_OLD_FORMAT" || pr.code === "URL_RETRIEVED_FROM") &&
            pr.before != null
        );
        if (fixableProblems.length === 0) {
          passed++;
          continue;
        }
        const p = model.paragraphs[ref.paragraphIndex]!;
        for (const problem of fixableProblems) {
          if (doFix && replaceParagraphText(p.el, problem.before!, problem.after ?? "")) {
            markDocDirty(ctx);
            fixedCount++;
            ctx.addChange({
              ruleId: "APA-REFERENCE-004",
              category: "references",
              location: loc(p),
              before: problem.before!,
              after: problem.after || "(removed)",
              reason: problem.message,
              confidence: 0.95,
            });
          } else {
            worst = "warning";
            ctx.addIssue({
              ruleId: "APA-REFERENCE-004",
              category: "references",
              severity: "warning",
              status: "warning",
              message: problem.message,
              location: loc(p),
              originalValue: problem.before,
              suggestedValue: problem.after,
              confidence: 0.9,
              autoFixable: true,
              userResolutionRequired: false,
            });
          }
        }
      }
      return result("APA-REFERENCE-004", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-REFERENCE-005",
    category: "references",
    description: "Every reference entry is cited in the text.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.references.length > 0,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "user_review" | null = null;
      for (const usage of analysis.matching.referenceUsage) {
        checked++;
        if (usage.cited) {
          passed++;
          continue;
        }
        const ref = analysis.references[usage.referenceIndex]!;
        const p = model.paragraphs[ref.paragraphIndex]!;
        worst = "user_review";
        ctx.addIssue({
          ruleId: "APA-REFERENCE-005",
          category: "references",
          severity: "warning",
          status: "user_review",
          message: `Reference "${excerptOf(ref.raw, 70)}" does not appear to be cited in the text.`,
          explanation:
            "APA 7 requires every reference to be cited in the body (annotated bibliographies excepted). Cite it, remove it, or confirm it is intentional.",
          location: loc(p),
          confidence: 0.8,
          autoFixable: false,
          userResolutionRequired: true,
          resolutionOptions: [
            { id: "will_cite", label: "I will cite it", description: "I will add a citation in the body." },
            { id: "will_remove", label: "I will remove it", description: "I will delete this reference." },
            { id: "intentional", label: "Intentional", description: "This entry is intentionally uncited (e.g. general reference)." },
          ],
        });
      }
      return result("APA-REFERENCE-005", checked, passed, false, worst);
    },
  },

  {
    id: "APA-REFERENCE-006",
    category: "references",
    description: "Minimum reference count (instructor requirement).",
    severity: "warning",
    applies: (ctx) => ctx.req.minReferences != null,
    run(ctx) {
      const min = ctx.req.minReferences!;
      const count = ctx.analysis.references.length;
      if (count >= min) return result("APA-REFERENCE-006", 1, 1, false, null);
      ctx.addIssue({
        ruleId: "APA-REFERENCE-006",
        category: "references",
        severity: "warning",
        status: "fail",
        message: `Instructor requires at least ${min} references; ${count} found.`,
        confidence: 0.9,
        autoFixable: false,
        userResolutionRequired: false,
      });
      return result("APA-REFERENCE-006", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-REFERENCE-007",
    category: "references",
    description: "Reference entry structure (authors, year, ampersand).",
    severity: "info",
    applies: (ctx) => ctx.analysis.references.length > 0,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      for (const ref of analysis.references) {
        checked++;
        const advisory = ref.problems.filter(
          (pr) => pr.code === "MISSING_YEAR" || pr.code === "AMPERSAND_MISSING"
        );
        if (advisory.length === 0) {
          passed++;
          continue;
        }
        const p = model.paragraphs[ref.paragraphIndex]!;
        for (const problem of advisory) {
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-REFERENCE-007",
            category: "references",
            severity: "info",
            status: "warning",
            message: `${problem.message} — "${excerptOf(ref.raw, 60)}"`,
            location: loc(p),
            confidence: 0.75,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-REFERENCE-007", checked, passed, false, worst);
    },
  },
];
