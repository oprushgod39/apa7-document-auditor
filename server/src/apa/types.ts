import type { DocumentModel } from "../docx/model.js";

export type PaperType = "student" | "professional";
export type ProcessingMode = "check" | "format" | "format_verify";

export type Severity = "error" | "warning" | "info";
export type Confidence = "high" | "medium" | "low";

/** Final status of a rule/issue in the compliance report. */
export type IssueStatus =
  | "pass"
  | "fixed"
  | "warning"
  | "user_review"
  | "fail"
  | "not_applicable"
  | "unverified";

export type RuleCategory =
  | "layout"
  | "paragraph"
  | "title_page"
  | "headings"
  | "abstract"
  | "citations"
  | "references"
  | "quotations"
  | "tables"
  | "figures"
  | "footnotes"
  | "statistics"
  | "numbers"
  | "abbreviations";

export interface IssueLocation {
  paragraphIndex?: number;
  blockIndex?: number;
  tableIndex?: number;
  description?: string;
  /** Short excerpt of the paragraph for user orientation. Truncated. */
  excerpt?: string;
}

export interface Issue {
  id: string;
  ruleId: string;
  category: RuleCategory;
  severity: Severity;
  status: IssueStatus;
  message: string;
  explanation?: string;
  location?: IssueLocation;
  originalValue?: string;
  suggestedValue?: string;
  confidence: number; // 0..1
  autoFixable: boolean;
  userResolutionRequired: boolean;
  /** Options the user may pick from when resolving (e.g. heading levels). */
  resolutionOptions?: ResolutionOption[];
  /** Set once the user resolves the issue. */
  resolution?: { optionId: string; note?: string };
}

export interface ResolutionOption {
  id: string;
  label: string;
  description?: string;
}

export interface Change {
  id: string;
  ruleId: string;
  category: RuleCategory;
  location: IssueLocation;
  before: string;
  after: string;
  reason: string;
  confidence: number;
  stage: "format" | "resolution";
  timestamp: string;
  /** User can exclude a change before generating the output. */
  excluded?: boolean;
  /** True when the change cannot be individually excluded (document-wide). */
  documentWide?: boolean;
}

export interface TitlePageMetadata {
  title?: string;
  author?: string;
  institution?: string;
  courseNumber?: string;
  courseName?: string;
  instructor?: string;
  dueDate?: string;
  authorNote?: string;
  runningHead?: string;
}

/** Instructor/assignment-level overrides layered on the APA 7 baseline. */
export interface InstructorRequirements {
  rawText?: string;
  font?: string;
  fontSizePt?: number;
  abstractRequired?: boolean | null;
  runningHeadRequired?: boolean | null;
  minReferences?: number | null;
  /** Free-form notes we could not interpret automatically. */
  uninterpreted: string[];
}

export interface DocumentSettings {
  paperType: PaperType;
  mode: ProcessingMode;
  preserveWording: boolean;
  fixCitationMechanics: boolean;
  verifyMetadata: boolean;
  metadata: TitlePageMetadata;
  instructor: InstructorRequirements;
}

/** Effective formatting requirements after instructor overrides. */
export interface EffectiveRequirements {
  font: string;
  fontSizePt: number;
  approvedFonts: { name: string; sizePt: number }[];
  marginTwips: number;
  lineSpacing: number; // w:line with lineRule auto: 480 = double
  firstLineIndentTwips: number; // 720 = 0.5"
  hangingIndentTwips: number;
  abstractRequired: boolean | null; // null = only if present
  runningHeadRequired: boolean;
  minReferences: number | null;
  overrides: string[]; // human-readable "Instructor override" notes
}

// --- Rule engine -----------------------------------------------------------

export interface RuleContext {
  model: DocumentModel;
  settings: DocumentSettings;
  req: EffectiveRequirements;
  analysis: import("./analysis.js").DocumentAnalysis;
  /** True during the audit pass — rules must not mutate. */
  auditOnly: boolean;
  addIssue(issue: Omit<Issue, "id">): Issue;
  addChange(change: Omit<Change, "id" | "timestamp" | "stage">): Change;
  nextId(prefix: string): string;
}

export interface RuleResult {
  ruleId: string;
  status: IssueStatus;
  checked: number; // how many items examined
  passed: number;
}

export interface ApaRule {
  id: string;
  category: RuleCategory;
  description: string;
  severity: Severity;
  /** Whether the rule applies to this document/settings. */
  applies(ctx: RuleContext): boolean;
  /**
   * Check the document, reporting issues via ctx.addIssue. When
   * `fix` is true and the rule is safely auto-fixable, apply corrections
   * (recording them via ctx.addChange) and report status "fixed".
   */
  run(ctx: RuleContext, fix: boolean): RuleResult | Promise<RuleResult>;
}

export const TWIPS_PER_INCH = 1440;
export const APA_DOUBLE_LINE = 480;
export const APA_FIRST_LINE_INDENT = 720;
export const APA_HANGING_INDENT = 720;

export const DEFAULT_APPROVED_FONTS: { name: string; sizePt: number }[] = [
  { name: "Times New Roman", sizePt: 12 },
  { name: "Calibri", sizePt: 11 },
  { name: "Arial", sizePt: 11 },
  { name: "Lucida Sans Unicode", sizePt: 10 },
  { name: "Georgia", sizePt: 11 },
  { name: "Computer Modern", sizePt: 10 },
];

export function confidenceScore(c: Confidence): number {
  return c === "high" ? 0.95 : c === "medium" ? 0.7 : 0.4;
}

export function excerptOf(text: string, max = 80): string {
  const t = text.trim().replace(/\s+/g, " ");
  return t.length <= max ? t : t.slice(0, max - 1) + "…";
}
