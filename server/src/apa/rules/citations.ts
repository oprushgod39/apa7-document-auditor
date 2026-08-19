import type { ApaRule } from "../types.js";
import { replaceParagraphText } from "../../docx/text.js";
import { result, loc, markDocDirty } from "./util.js";
import { excerptOf } from "../types.js";

/** F/G. In-text citation rules and citation ↔ reference matching. */
export const citationRules: ApaRule[] = [
  {
    id: "APA-CITATION-001",
    category: "citations",
    description:
      "Citation mechanics: & vs and, et al. punctuation, comma before year.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.citations.length > 0,
    run(ctx, fix) {
      const { model, analysis, settings } = ctx;
      const doFix = fix && settings.fixCitationMechanics;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "warning" | null = null;

      for (const cit of analysis.citations) {
        checked++;
        if (cit.problems.length === 0) {
          passed++;
          continue;
        }
        const p = model.paragraphs[cit.paragraphIndex];
        if (!p) continue;
        for (const problem of cit.problems) {
          // Deterministic, single-interpretation fixes only.
          const replacement = deterministicFix(cit.raw, problem.code);
          if (doFix && replacement && replacement !== cit.raw) {
            const ok = replaceParagraphText(p.el, cit.raw, replacement);
            if (ok) {
              markDocDirty(ctx);
              fixedCount++;
              ctx.addChange({
                ruleId: "APA-CITATION-001",
                category: "citations",
                location: loc(p),
                before: cit.raw,
                after: replacement,
                reason: problem.message,
                confidence: 0.92,
              });
              continue;
            }
          }
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-CITATION-001",
            category: "citations",
            severity: "warning",
            status: "warning",
            message: `${problem.message} — ${cit.raw}`,
            location: loc(p),
            originalValue: cit.raw,
            suggestedValue: replacement ?? undefined,
            confidence: 0.85,
            autoFixable: replacement != null,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-CITATION-001", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-CITATION-008",
    category: "citations",
    description: "Every in-text citation has a matching reference entry.",
    severity: "error",
    applies: (ctx) => ctx.analysis.citations.length > 0,
    run(ctx) {
      const { model, analysis } = ctx;
      const hasRefList = analysis.referencesHeadingIndex != null;
      let checked = 0;
      let passed = 0;
      let worst: "user_review" | "warning" | null = null;

      for (const match of analysis.matching.citationMatches) {
        checked++;
        const cit = analysis.citations[match.citationIndex]!;
        if (match.status === "exact" || match.status === "probable") {
          passed++;
          continue;
        }
        const p = model.paragraphs[cit.paragraphIndex];
        if (match.status === "missing_reference") {
          worst = "user_review";
          ctx.addIssue({
            ruleId: "APA-CITATION-008",
            category: "citations",
            severity: "error",
            status: "user_review",
            message: hasRefList
              ? `Citation ${cit.raw} has no matching reference-list entry.`
              : `Citation ${cit.raw} found, but no References section was detected.`,
            explanation:
              "Add the missing reference to the reference list, or confirm the citation refers to a personal communication (which is not listed in references).",
            location: p ? loc(p) : undefined,
            originalValue: cit.raw,
            confidence: 0.9,
            autoFixable: false,
            userResolutionRequired: true,
            resolutionOptions: [
              { id: "will_add", label: "I will add the reference", description: "I will add this entry to the reference list." },
              { id: "personal_communication", label: "Personal communication", description: "This source is a personal communication and is correctly not in the reference list." },
              { id: "not_a_citation", label: "Not a citation", description: "This text is not actually a citation." },
            ],
          });
        } else {
          if (worst !== "user_review") worst = "warning";
          ctx.addIssue({
            ruleId: "APA-CITATION-008",
            category: "citations",
            severity: "warning",
            status: "warning",
            message: `Citation ${cit.raw} matches multiple reference entries (ambiguous).`,
            explanation: match.detail,
            location: p ? loc(p) : undefined,
            confidence: 0.75,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-CITATION-008", checked, passed, false, worst);
    },
  },

  {
    id: "APA-CITATION-002",
    category: "citations",
    description: "Direct quotations include a page or paragraph locator.",
    severity: "info",
    applies: () => true,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      const refSet = new Set(analysis.referenceEntryIndexes);

      for (const p of model.paragraphs) {
        if (p.isEmpty || refSet.has(p.index) || p.index < analysis.bodyStartIndex) continue;
        if (analysis.referencesHeadingIndex != null && p.index >= analysis.referencesHeadingIndex) continue;
        // Quoted spans of 6+ words are treated as likely direct quotations.
        const quotes = [...p.text.matchAll(/[“"]([^”"]{20,})[”"]/g)];
        for (const q of quotes) {
          if (q[1]!.split(/\s+/).length < 6) continue;
          checked++;
          const citationsHere = analysis.citations.filter(
            (c) => c.paragraphIndex === p.index
          );
          const hasLocator = citationsHere.some((c) => c.locator != null);
          if (hasLocator) {
            passed++;
          } else {
            worst = "warning";
            ctx.addIssue({
              ruleId: "APA-CITATION-002",
              category: "citations",
              severity: "info",
              status: "warning",
              message: `A direct quotation may be missing a page/paragraph locator: “${excerptOf(q[1]!, 60)}”`,
              explanation:
                "APA 7 requires a locator (p., pp., para.) for direct quotations from paginated sources. Sources without page numbers may use paragraph numbers, headings, or timestamps — so this is advisory.",
              location: loc(p),
              confidence: 0.6,
              autoFixable: false,
              userResolutionRequired: false,
            });
          }
        }
      }
      return result("APA-CITATION-002", checked, passed, false, worst);
    },
  },
];

function deterministicFix(raw: string, code: string): string | null {
  switch (code) {
    case "AND_IN_PARENTHETICAL":
      // (Smith and Patel, 2024) → (Smith & Patel, 2024): unambiguous.
      return raw.replace(/\b and \b/g, " & ").replace(/\band\b(?=\s+[A-ZÀ-Þ])/g, "&");
    case "AMPERSAND_IN_NARRATIVE":
      return raw.replace(/\s*&\s*/g, " and ");
    case "ET_AL_MISSING_PERIOD":
      return raw.replace(/\bet al(?!\.)/g, "et al.");
    default:
      return null; // MISSING_COMMA_BEFORE_YEAR etc. can be ambiguous — advise only
  }
}
