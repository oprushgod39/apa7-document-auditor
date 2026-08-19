import type { EngineRun } from "../apa/engine.js";
import type {
  Change,
  Issue,
  IssueStatus,
  RuleCategory,
} from "../apa/types.js";
import type { VerificationResult } from "../verify/provider.js";

/**
 * Independent APA compliance auditor.
 *
 * The audit runs the full rule engine in check-only mode against the
 * MODIFIED document model (re-parsed from the mutated package) — it never
 * assumes a rule passed just because a fix was applied.
 */

export const CATEGORY_LABELS: Record<RuleCategory, string> = {
  layout: "Layout",
  paragraph: "Paragraphs",
  title_page: "Title Page",
  headings: "Headings",
  abstract: "Abstract",
  citations: "Citations",
  references: "References",
  quotations: "Quotations",
  tables: "Tables",
  figures: "Figures",
  footnotes: "Footnotes",
  statistics: "Statistics",
  numbers: "Numbers",
  abbreviations: "Abbreviations",
};

export interface CategorySummary {
  category: RuleCategory;
  label: string;
  status: IssueStatus;
  rulesChecked: number;
  rulesPassed: number;
  issueCount: number;
}

export interface VerificationSummary {
  provider: string;
  attempted: number;
  verified: number;
  probable: number;
  mismatched: number;
  unverified: number;
  providerUnavailable: boolean;
  results: VerificationResult[];
}

export interface ComplianceReport {
  /** Primary completion state — never a misleading single score. */
  state: "apa_validated" | "review_required";
  unresolvedCount: number;
  /** Supplementary percentage of applicable rules resolved. */
  rulesResolvedPercent: number;
  /** Count of rules applicable to this document (for recomputation). */
  applicableRules: number;
  categories: CategorySummary[];
  issues: IssueWithResolution[];
  changes: Change[];
  changesApplied: number;
  verification: VerificationSummary | null;
  instructorOverrides: string[];
  generatedAt: string;
}

export interface IssueWithResolution extends Issue {
  resolved: boolean;
}

/** Stable identity for an issue across engine runs. */
export function issueKey(issue: Issue): string {
  return [
    issue.ruleId,
    issue.location?.paragraphIndex ?? issue.location?.tableIndex ?? "",
    issue.originalValue ?? issue.message,
  ].join("|");
}

const STATUS_RANK: Record<IssueStatus, number> = {
  fail: 6,
  user_review: 5,
  warning: 4,
  unverified: 3,
  fixed: 2,
  pass: 1,
  not_applicable: 0,
};

export function buildReport(
  auditRun: EngineRun,
  allChanges: Change[],
  verification: VerificationResult[] | null,
  resolutions: Map<string, { optionId: string; note?: string }>,
  ruleCategoryById: Map<string, RuleCategory>
): ComplianceReport {
  // Attach stored user resolutions to audit issues.
  const issues: IssueWithResolution[] = auditRun.issues.map((issue) => {
    const res = resolutions.get(issueKey(issue));
    return {
      ...issue,
      resolution: res,
      resolved: res != null || !issue.userResolutionRequired,
    };
  });

  // Unresolved = issues requiring user resolution without one, plus errors.
  const unresolved = issues.filter(
    (i) =>
      (i.userResolutionRequired && i.resolution == null) ||
      (i.severity === "error" && i.status === "fail")
  );

  // Category summaries from audit rule results.
  const byCategory = new Map<RuleCategory, CategorySummary>();
  for (const rr of auditRun.ruleResults) {
    const category = ruleCategoryById.get(rr.ruleId);
    if (!category) continue;
    let entry = byCategory.get(category);
    if (!entry) {
      entry = {
        category,
        label: CATEGORY_LABELS[category],
        status: "not_applicable",
        rulesChecked: 0,
        rulesPassed: 0,
        issueCount: 0,
      };
      byCategory.set(category, entry);
    }
    if (rr.status !== "not_applicable") {
      entry.rulesChecked++;
      let effective: IssueStatus = rr.status;
      // A rule whose only outstanding items are user-resolved counts as pass.
      if (effective === "user_review") {
        const open = unresolved.some(
          (i) => i.ruleId === rr.ruleId && i.userResolutionRequired && i.resolution == null
        );
        if (!open) effective = "pass";
      }
      if (effective === "pass" || effective === "fixed") entry.rulesPassed++;
      if (STATUS_RANK[effective] > STATUS_RANK[entry.status]) {
        entry.status = effective;
      }
    }
  }
  for (const issue of issues) {
    const entry = byCategory.get(issue.category);
    if (entry) entry.issueCount++;
  }

  const categories = [...byCategory.values()].sort(
    (a, b) => STATUS_RANK[b.status] - STATUS_RANK[a.status]
  );

  const applicable = auditRun.ruleResults.filter((r) => r.status !== "not_applicable");
  const resolvedRules = applicable.filter((r) => {
    if (r.status === "pass" || r.status === "fixed") return true;
    if (r.status === "user_review" || r.status === "warning" || r.status === "fail") {
      // resolved if no open unresolved issue remains for this rule
      return !unresolved.some((i) => i.ruleId === r.ruleId);
    }
    return false;
  });

  let verificationSummary: VerificationSummary | null = null;
  if (verification) {
    verificationSummary = {
      provider: verification[0]?.provider ?? "none",
      attempted: verification.length,
      verified: verification.filter((v) => v.status === "verified").length,
      probable: verification.filter((v) => v.status === "probable").length,
      mismatched: verification.filter((v) => v.status === "mismatch").length,
      unverified: verification.filter((v) => v.status === "unverified").length,
      providerUnavailable: verification.some(
        (v) => v.status === "provider_unavailable"
      ),
      results: verification,
    };
  }

  return {
    state: unresolved.length === 0 ? "apa_validated" : "review_required",
    unresolvedCount: unresolved.length,
    rulesResolvedPercent:
      applicable.length === 0
        ? 100
        : Math.round((resolvedRules.length / applicable.length) * 100),
    applicableRules: applicable.length,
    categories,
    issues,
    changes: allChanges,
    changesApplied: allChanges.filter((c) => !c.excluded).length,
    verification: verificationSummary,
    instructorOverrides: auditRun.req.overrides,
    generatedAt: new Date().toISOString(),
  };
}
