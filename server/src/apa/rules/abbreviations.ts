import type { ApaRule } from "../types.js";
import { result } from "./util.js";

/**
 * S. Abbreviation rules — ADVISORY ONLY. Sentences are never rewritten.
 */
export const abbreviationRules: ApaRule[] = [
  {
    id: "APA-ABBR-001",
    category: "abbreviations",
    description:
      "Abbreviations are introduced with their full term at first use.",
    severity: "info",
    applies: () => true,
    run(ctx) {
      const { model, analysis } = ctx;
      const refSet = new Set(analysis.referenceEntryIndexes);
      const bodyText = model.paragraphs
        .filter(
          (p) =>
            !p.isEmpty &&
            !p.insideTable &&
            !refSet.has(p.index) &&
            p.index >= analysis.bodyStartIndex &&
            (analysis.referencesHeadingIndex == null ||
              p.index < analysis.referencesHeadingIndex)
        )
        .map((p) => p.text)
        .join("\n");

      const IGNORE = new Set([
        "APA", "USA", "US", "UK", "AM", "PM", "AD", "BC", "IQ", "AIDS", "HIV",
        "DNA", "RNA", "PDF", "HTML", "URL", "DOI", "ET", "AL", "II", "III", "IV",
      ]);
      // Count abbreviation usage and whether a "(ABBR)" definition exists.
      const counts = new Map<string, number>();
      for (const m of bodyText.matchAll(/\b([A-Z]{2,6})s?\b/g)) {
        const ab = m[1]!;
        if (IGNORE.has(ab)) continue;
        counts.set(ab, (counts.get(ab) ?? 0) + 1);
      }
      const defined = new Set<string>();
      const definitionCounts = new Map<string, number>();
      for (const m of bodyText.matchAll(/\(([A-Z]{2,6})s?\)/g)) {
        defined.add(m[1]!);
        definitionCounts.set(m[1]!, (definitionCounts.get(m[1]!) ?? 0) + 1);
      }

      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      for (const [ab, count] of counts) {
        if (count < 3) continue; // only recurring abbreviations matter
        checked++;
        if (defined.has(ab)) {
          const defs = definitionCounts.get(ab) ?? 0;
          if (defs > 1) {
            worst = "warning";
            ctx.addIssue({
              ruleId: "APA-ABBR-001",
              category: "abbreviations",
              severity: "info",
              status: "warning",
              message: `"${ab}" appears to be defined ${defs} times. Define an abbreviation only at first use.`,
              explanation: "Advisory only — wording is never changed automatically.",
              confidence: 0.6,
              autoFixable: false,
              userResolutionRequired: false,
            });
          } else {
            passed++;
          }
        } else {
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-ABBR-001",
            category: "abbreviations",
            severity: "info",
            status: "warning",
            message: `"${ab}" is used ${count} times but no "(${ab})" definition was found.`,
            explanation:
              "APA introduces abbreviations with the full term followed by the abbreviation in parentheses at first use. Advisory only.",
            confidence: 0.6,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-ABBR-001", Math.max(checked, 1), checked === 0 ? 1 : passed, false, worst);
    },
  },
];
