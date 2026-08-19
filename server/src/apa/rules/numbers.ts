import type { ApaRule } from "../types.js";
import { result, loc } from "./util.js";

/**
 * T. Numbers and measurements — ADVISORY ONLY. Wording is never changed.
 */
export const numberRules: ApaRule[] = [
  {
    id: "APA-NUM-001",
    category: "numbers",
    description: "Sentences should not begin with a numeral.",
    severity: "info",
    applies: () => true,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      const refSet = new Set(analysis.referenceEntryIndexes);
      for (const p of model.paragraphs) {
        if (p.isEmpty || p.insideTable || refSet.has(p.index) || p.hasNumbering) continue;
        if (p.index < analysis.bodyStartIndex) continue;
        // Sentence-initial numerals (start of paragraph or after ". ")
        const matches = [
          ...p.text.matchAll(/(?:^|[.!?]\s+)(\d[\d,.]*)\s+[a-z]/g),
        ];
        for (const m of matches.slice(0, 2)) {
          // Years/dates at sentence start are commonly acceptable in context.
          if (/^(?:1[6-9]|20)\d{2}$/.test(m[1]!)) continue;
          checked++;
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-NUM-001",
            category: "numbers",
            severity: "info",
            status: "warning",
            message: `A sentence appears to begin with the numeral "${m[1]}". APA spells out numbers that begin a sentence.`,
            explanation:
              "Advisory only — your wording is never changed automatically.",
            location: loc(p),
            confidence: 0.6,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-NUM-001", Math.max(checked, 1), checked === 0 ? 1 : passed, false, worst);
    },
  },

  {
    id: "APA-NUM-002",
    category: "numbers",
    description: "Use numerals with % and units of measurement.",
    severity: "info",
    applies: () => true,
    run(ctx) {
      const { model, analysis } = ctx;
      let flagged = 0;
      let worst: "warning" | null = null;
      const words =
        /\b(one|two|three|four|five|six|seven|eight|nine|ten)\s+(percent|kg|km|cm|mm|ms|mg|ml)\b/gi;
      for (const p of model.paragraphs) {
        if (p.isEmpty || p.insideTable || p.index < analysis.bodyStartIndex) continue;
        for (const m of p.text.matchAll(words)) {
          flagged++;
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-NUM-002",
            category: "numbers",
            severity: "info",
            status: "warning",
            message: `"${m[0]}" — APA uses numerals with percentages and measurement units (e.g. "5%", "3 kg").`,
            explanation: "Advisory only — wording is never changed automatically.",
            location: loc(p),
            confidence: 0.7,
            autoFixable: false,
            userResolutionRequired: false,
          });
          if (flagged >= 5) break;
        }
        if (flagged >= 5) break;
      }
      return result("APA-NUM-002", Math.max(flagged, 1), flagged === 0 ? 1 : 0, false, worst);
    },
  },
];
