/** Typed API client for the APA 7 Document Auditor backend. */

export interface DetectedInfo {
  metadata: {
    title?: string;
    author?: string;
    institution?: string;
    courseNumber?: string;
    courseName?: string;
    instructor?: string;
    dueDate?: string;
  };
  hasTitlePage: boolean;
  hasAbstract: boolean;
  headings: number;
  citations: number;
  references: number;
  tables: number;
  images: number;
  paragraphs: number;
  footnotes: number;
}

export interface Stage {
  key: string;
  label: string;
  status: "pending" | "running" | "done" | "skipped" | "failed";
}

export interface SessionInfo {
  id: string;
  originalName: string;
  status: "uploaded" | "processing" | "ready" | "error";
  stages: Stage[];
  error: string | null;
  settings: {
    paperType: string;
    mode: string;
    preserveWording: boolean;
    fixCitationMechanics: boolean;
    verifyMetadata: boolean;
  };
}

export interface UploadResponse extends SessionInfo {
  detected: DetectedInfo;
}

export interface ResolutionOption {
  id: string;
  label: string;
  description?: string;
}

export interface Issue {
  id: string;
  ruleId: string;
  category: string;
  severity: "error" | "warning" | "info";
  status: string;
  message: string;
  explanation?: string;
  location?: {
    paragraphIndex?: number;
    tableIndex?: number;
    description?: string;
    excerpt?: string;
  };
  originalValue?: string;
  suggestedValue?: string;
  confidence: number;
  autoFixable: boolean;
  userResolutionRequired: boolean;
  resolutionOptions?: ResolutionOption[];
  resolution?: { optionId: string; note?: string };
  resolved: boolean;
}

export interface Change {
  id: string;
  ruleId: string;
  category: string;
  location: { paragraphIndex?: number; description?: string; excerpt?: string };
  before: string;
  after: string;
  reason: string;
  confidence: number;
  stage: string;
  documentWide?: boolean;
}

export interface CategorySummary {
  category: string;
  label: string;
  status: string;
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
}

export interface Report {
  state: "apa_validated" | "review_required";
  unresolvedCount: number;
  rulesResolvedPercent: number;
  categories: CategorySummary[];
  issues: Issue[];
  changes: Change[];
  changesApplied: number;
  verification: VerificationSummary | null;
  instructorOverrides: string[];
  generatedAt: string;
}

export interface OutlineEntry {
  index: number;
  kind: string;
  level?: number;
  confidence?: string;
  text: string;
  issues: number;
}

export interface ReportResponse extends SessionInfo {
  report: Report;
  outline: OutlineEntry[];
  instructorUninterpreted: string[];
}

export interface ProcessSettings {
  paperType: "student" | "professional";
  mode: "check" | "format" | "format_verify";
  preserveWording: boolean;
  fixCitationMechanics: boolean;
  verifyMetadata: boolean;
  metadata: Record<string, string>;
  instructorRequirements: string;
}

export class ApiError extends Error {
  code: string;
  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

async function handle<T>(res: Response): Promise<T> {
  if (res.ok) return (await res.json()) as T;
  let code = "INTERNAL";
  let message = "Something went wrong.";
  try {
    const body = await res.json();
    code = body?.error?.code ?? code;
    message = body?.error?.message ?? message;
  } catch {
    /* non-JSON error */
  }
  throw new ApiError(code, message);
}

export async function uploadDocument(file: File): Promise<UploadResponse> {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch("/api/documents", { method: "POST", body: form });
  return handle<UploadResponse>(res);
}

export async function startProcessing(
  id: string,
  settings: ProcessSettings
): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/process`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(settings),
  });
  return handle<SessionInfo>(res);
}

export async function getStatus(id: string): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/status`);
  return handle<SessionInfo>(res);
}

export async function getReport(id: string): Promise<ReportResponse> {
  const res = await fetch(`/api/documents/${id}/report`);
  return handle<ReportResponse>(res);
}

export async function resolveIssue(
  id: string,
  issueKey: string,
  optionId: string
): Promise<{ report: Report }> {
  const res = await fetch(`/api/documents/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ issueKey, optionId }),
  });
  return handle<{ report: Report }>(res);
}

export async function regenerate(id: string): Promise<SessionInfo> {
  const res = await fetch(`/api/documents/${id}/generate`, { method: "POST" });
  return handle<SessionInfo>(res);
}

export function issueKeyOf(issue: Issue): string {
  return [
    issue.ruleId,
    issue.location?.paragraphIndex ?? issue.location?.tableIndex ?? "",
    issue.originalValue ?? issue.message,
  ].join("|");
}

export const downloadUrl = (id: string) => `/api/documents/${id}/download`;
export const reportDownloadUrl = (id: string) => `/api/documents/${id}/report.html`;
