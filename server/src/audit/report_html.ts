import type { ComplianceReport } from "./auditor.js";
import { config } from "../config.js";

/** Standalone HTML compliance report for download. */
export function renderReportHtml(
  report: ComplianceReport,
  documentName: string
): string {
  const esc = (s: string) =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const statusLabel: Record<string, string> = {
    pass: "PASS",
    fixed: "FIXED",
    warning: "WARNING",
    user_review: "USER REVIEW",
    fail: "FAIL",
    not_applicable: "NOT APPLICABLE",
    unverified: "UNVERIFIED",
  };
  const unresolved = report.issues.filter(
    (i) => i.userResolutionRequired && i.resolution == null
  );
  const rows = report.categories
    .map(
      (c) => `<tr><td>${esc(c.label)}</td><td class="s-${c.status}">${statusLabel[c.status] ?? c.status}</td><td>${c.rulesPassed}/${c.rulesChecked}</td><td>${c.issueCount}</td></tr>`
    )
    .join("");
  const issueRows = report.issues
    .map(
      (i) =>
        `<tr><td>${esc(i.ruleId)}</td><td class="s-${i.status}">${statusLabel[i.status] ?? i.status}${i.resolved && i.userResolutionRequired ? " (resolved)" : ""}</td><td>${esc(i.message)}${i.location?.excerpt ? `<br><small>near: “${esc(i.location.excerpt)}”</small>` : ""}</td></tr>`
    )
    .join("");
  const changeRows = report.changes
    .map(
      (c) =>
        `<tr><td>${esc(c.ruleId)}</td><td>${esc(describeLocation(c.location))}</td><td>${esc(c.before)}</td><td>${esc(c.after)}</td></tr>`
    )
    .join("");
  const verification = report.verification
    ? `<p>Provider: ${esc(report.verification.provider)} — verified ${report.verification.verified}, probable ${report.verification.probable}, mismatched ${report.verification.mismatched}, unverified ${report.verification.unverified}${report.verification.providerUnavailable ? " (provider temporarily unavailable for some entries)" : ""}.</p>`
    : "<p>External metadata verification was not run.</p>";

  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>APA 7 Compliance Report — ${esc(documentName)}</title>
<style>
body{font-family:Georgia,'Times New Roman',serif;max-width:60rem;margin:2rem auto;padding:0 1rem;color:#1a1a1a;line-height:1.5}
h1{font-size:1.6rem;border-bottom:2px solid #1a1a1a;padding-bottom:.4rem}
h2{font-size:1.15rem;margin-top:2rem}
table{border-collapse:collapse;width:100%;font-size:.9rem}
th,td{border:1px solid #ccc;padding:.4rem .6rem;text-align:left;vertical-align:top}
th{background:#f4f2ee}
.badge{display:inline-block;padding:.25rem .75rem;border-radius:4px;font-weight:bold}
.validated{background:#e8f2e8;color:#1e5b1e;border:1px solid #1e5b1e}
.review{background:#fdf3e0;color:#8a5a00;border:1px solid #8a5a00}
.s-pass,.s-fixed{color:#1e5b1e}.s-warning,.s-user_review{color:#8a5a00}.s-fail{color:#9b1c1c}.s-unverified,.s-not_applicable{color:#666}
small{color:#555}
</style></head><body>
<h1>APA 7 Compliance Report</h1>
<p><strong>${esc(config.productName)}</strong> — ${esc(documentName)}<br>
Generated ${esc(new Date(report.generatedAt).toLocaleString())}</p>
<p><span class="badge ${report.state === "apa_validated" ? "validated" : "review"}">
${report.state === "apa_validated" ? "APA 7 VALIDATED — all applicable rules have been checked or resolved" : `REVIEW REQUIRED — ${report.unresolvedCount} item(s) need your attention`}
</span></p>
<p>Rules resolved: ${report.rulesResolvedPercent}% &nbsp;·&nbsp; Formatting changes applied: ${report.changesApplied} &nbsp;·&nbsp; Unresolved issues: ${unresolved.length}</p>
${report.instructorOverrides.length > 0 ? `<p><em>${report.instructorOverrides.map(esc).join("<br>")}</em></p>` : ""}
<h2>Category summary</h2>
<table><tr><th>Category</th><th>Status</th><th>Rules passed</th><th>Findings</th></tr>${rows}</table>
<h2>Reference verification</h2>
${verification}
<h2>Issues (${report.issues.length})</h2>
${report.issues.length ? `<table><tr><th>Rule</th><th>Status</th><th>Detail</th></tr>${issueRows}</table>` : "<p>No issues.</p>"}
<h2>Applied changes (${report.changes.length})</h2>
${report.changes.length ? `<table><tr><th>Rule</th><th>Location</th><th>Before</th><th>After</th></tr>${changeRows}</table>` : "<p>No changes were applied.</p>"}
<p><small>APA Validated means every applicable rule has either been automatically verified or explicitly resolved by you. It is not a guarantee of publisher- or instructor-specific acceptance.</small></p>
</body></html>`;
}

function describeLocation(loc: { paragraphIndex?: number; description?: string; excerpt?: string }): string {
  if (loc.description) return loc.description;
  if (loc.paragraphIndex != null) {
    return `Paragraph ${loc.paragraphIndex + 1}${loc.excerpt ? ` (“${loc.excerpt}”)` : ""}`;
  }
  return "";
}
