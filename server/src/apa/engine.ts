import { randomUUID } from "node:crypto";
import type { DocumentModel } from "../docx/model.js";
import { analyzeDocument, type DocumentAnalysis } from "./analysis.js";
import type {
  ApaRule,
  Change,
  DocumentSettings,
  EffectiveRequirements,
  Issue,
  RuleContext,
  RuleResult,
} from "./types.js";
import { effectiveRequirements } from "./requirements.js";
import { layoutRules } from "./rules/layout.js";
import { paragraphRules } from "./rules/paragraphs.js";
import { titlePageRules } from "./rules/title_page.js";
import { headingRules } from "./rules/headings.js";
import { abstractRules } from "./rules/abstract.js";
import { citationRules } from "./rules/citations.js";
import { referenceRules } from "./rules/references.js";
import { quotationRules } from "./rules/quotations.js";
import { tableRules } from "./rules/tables.js";
import { figureRules } from "./rules/figures.js";
import { statisticsRules } from "./rules/statistics.js";
import { numberRules } from "./rules/numbers.js";
import { abbreviationRules } from "./rules/abbreviations.js";

/** Rule registry. Institutional profiles can filter/replace this later. */
export function allRules(): ApaRule[] {
  return [
    ...layoutRules,
    ...paragraphRules,
    ...titlePageRules,
    ...headingRules,
    ...abstractRules,
    ...citationRules,
    ...referenceRules,
    ...quotationRules,
    ...tableRules,
    ...figureRules,
    ...statisticsRules,
    ...numberRules,
    ...abbreviationRules,
  ];
}

export interface EngineRun {
  issues: Issue[];
  changes: Change[];
  ruleResults: RuleResult[];
  analysis: DocumentAnalysis;
  req: EffectiveRequirements;
}

export interface EngineOptions {
  fix: boolean;
  auditOnly?: boolean;
  stage?: "format" | "resolution";
  /** Rule IDs disabled by configuration. */
  disabledRules?: Set<string>;
  /** Pre-computed analysis to reuse (avoids re-parsing). */
  analysis?: DocumentAnalysis;
}

export async function runEngine(
  model: DocumentModel,
  settings: DocumentSettings,
  opts: EngineOptions
): Promise<EngineRun> {
  const req = effectiveRequirements(settings);
  const analysis = opts.analysis ?? analyzeDocument(model);
  const issues: Issue[] = [];
  const changes: Change[] = [];
  const ruleResults: RuleResult[] = [];
  const stage = opts.stage ?? "format";

  const ctx: RuleContext = {
    model,
    settings,
    req,
    analysis,
    auditOnly: opts.auditOnly ?? false,
    addIssue(issue) {
      const full: Issue = { ...issue, id: randomUUID() };
      issues.push(full);
      return full;
    },
    addChange(change) {
      const full: Change = {
        ...change,
        id: randomUUID(),
        stage,
        timestamp: new Date().toISOString(),
      };
      changes.push(full);
      return full;
    },
    nextId: (prefix: string) => `${prefix}-${randomUUID().slice(0, 8)}`,
  };

  for (const rule of allRules()) {
    if (opts.disabledRules?.has(rule.id)) continue;
    let applies: boolean;
    try {
      applies = rule.applies(ctx);
    } catch {
      applies = false;
    }
    if (!applies) {
      ruleResults.push({ ruleId: rule.id, status: "not_applicable", checked: 0, passed: 0 });
      continue;
    }
    try {
      const result = await rule.run(ctx, opts.fix && !opts.auditOnly);
      ruleResults.push(result);
    } catch (err) {
      // A rule failure must never break the pipeline; report honestly.
      ruleResults.push({ ruleId: rule.id, status: "unverified", checked: 0, passed: 0 });
      ctx.addIssue({
        ruleId: rule.id,
        category: rule.category,
        severity: "warning",
        status: "unverified",
        message: `Rule ${rule.id} could not be evaluated for this document.`,
        explanation: err instanceof Error ? err.message : undefined,
        confidence: 1,
        autoFixable: false,
        userResolutionRequired: false,
      });
    }
  }

  return { issues, changes, ruleResults, analysis, req };
}
