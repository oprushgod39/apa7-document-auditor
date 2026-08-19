import type { ParagraphModel } from "../../docx/model.js";
import type { RuleContext, RuleResult, IssueStatus } from "../types.js";
import { excerptOf } from "../types.js";

export function result(
  ruleId: string,
  checked: number,
  passed: number,
  fixedAny: boolean,
  worstIssue: IssueStatus | null
): RuleResult {
  let status: IssueStatus;
  if (checked === 0) status = "not_applicable";
  else if (worstIssue && worstIssue !== "pass") status = worstIssue;
  else if (fixedAny) status = "fixed";
  else status = "pass";
  return { ruleId, status, checked, passed };
}

export function loc(p: ParagraphModel) {
  return {
    paragraphIndex: p.index,
    blockIndex: p.blockIndex,
    excerpt: excerptOf(p.text),
  };
}

/** Paragraphs in the main body (not title page, not inside tables). */
export function bodyParagraphs(ctx: RuleContext): ParagraphModel[] {
  const a = ctx.analysis;
  return ctx.model.paragraphs.filter(
    (p) => !p.insideTable && p.index >= a.titlePageEnd
  );
}

/** True if the paragraph is a classified heading. */
export function isHeadingParagraph(ctx: RuleContext, index: number): boolean {
  return ctx.analysis.headings.some(
    (h) => h.paragraphIndex === index && h.level > 0
  );
}

export function isReferenceEntry(ctx: RuleContext, index: number): boolean {
  return ctx.analysis.referenceEntryIndexes.includes(index);
}

export function markDocDirty(ctx: RuleContext): void {
  ctx.model.pkg.markDirty("word/document.xml");
}

export function markStylesDirty(ctx: RuleContext): void {
  if (ctx.model.pkg.has("word/styles.xml")) {
    ctx.model.pkg.markDirty("word/styles.xml");
  }
}
