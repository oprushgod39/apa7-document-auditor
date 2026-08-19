import type { ApaRule } from "../types.js";
import { result, loc } from "./util.js";
import { excerptOf } from "../types.js";

/**
 * R. Statistical expression rules — ADVISORY ONLY. Statistical text is never
 * modified automatically; interpretation depends on context.
 */
export const statisticsRules: ApaRule[] = [
  {
    id: "APA-STAT-001",
    category: "statistics",
    description:
      "Statistical expressions follow APA spacing and leading-zero rules.",
    severity: "info",
    applies: () => true,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      const refSet = new Set(analysis.referenceEntryIndexes);
      for (const p of model.paragraphs) {
        if (p.isEmpty || p.insideTable || refSet.has(p.index)) continue;
        if (p.index < analysis.bodyStartIndex) continue;
        const text = p.text;
        const findings: string[] = [];
        // p/r values cannot exceed 1 → no leading zero: p < 0.05 → p < .05
        for (const m of text.matchAll(/\b([pr])\s*([<>=≤≥])\s*(0\.\d+)/g)) {
          findings.push(
            `"${m[0]}" — ${m[1]} values cannot exceed 1, so APA omits the leading zero (${m[1]} ${m[2]} ${m[3]!.slice(1)}).`
          );
        }
        // Missing spaces around operators: p=.05
        for (const m of text.matchAll(/\b([ptrFMdn]|SD|SE|M)([<>=])(\.?\d)/g)) {
          findings.push(
            `"${m[0]}" — APA puts spaces around statistical operators (${m[1]} ${m[2]} ${m[3]}…).`
          );
        }
        if (findings.length === 0) continue;
        checked++;
        worst = "warning";
        for (const f of findings.slice(0, 3)) {
          ctx.addIssue({
            ruleId: "APA-STAT-001",
            category: "statistics",
            severity: "info",
            status: "warning",
            message: f,
            explanation:
              "Advisory only — statistical expressions are never modified automatically. Also check that statistical symbols (p, t, F, M, SD, r, d, n) are italicized.",
            location: loc(p),
            confidence: 0.75,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-STAT-001", Math.max(checked, 1), checked === 0 ? 1 : passed, false, worst);
    },
  },
];
