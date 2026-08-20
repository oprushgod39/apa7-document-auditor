import { useMemo, useState } from "react";
import {
  downloadUrl,
  issueKeyOf,
  regenerate,
  reportDownloadUrl,
  resolveIssue,
  type Issue,
  type ReportResponse,
  type UploadResponse,
} from "../lib/api";

type Filter = "all" | "errors" | "warnings" | "fixed" | "needs_review";

const STATUS_LABEL: Record<string, string> = {
  pass: "PASS",
  fixed: "FIXED",
  warning: "WARNING",
  user_review: "USER REVIEW",
  fail: "FAIL",
  not_applicable: "N/A",
  unverified: "UNVERIFIED",
};

export function ResultsScreen(props: {
  session: UploadResponse;
  data: ReportResponse;
  onReportUpdate: (data: ReportResponse) => void;
  onRegenerate: () => void;
  onReset: () => void;
}) {
  const { data, session } = props;
  const report = data.report;
  const [filter, setFilter] = useState<Filter>("all");
  const [showOutline, setShowOutline] = useState(false);
  const [showChanges, setShowChanges] = useState(false);
  const [resolving, setResolving] = useState<string | null>(null);
  const [headingResolved, setHeadingResolved] = useState(false);

  const issues = useMemo(() => {
    return report.issues.filter((i) => {
      switch (filter) {
        case "errors":
          return i.severity === "error";
        case "warnings":
          return i.severity === "warning" || i.severity === "info";
        case "fixed":
          return i.status === "fixed";
        case "needs_review":
          return i.userResolutionRequired && !i.resolution;
        default:
          return true;
      }
    });
  }, [report.issues, filter]);

  const resolve = async (issue: Issue, optionId: string) => {
    const key = issueKeyOf(issue);
    setResolving(key);
    try {
      const res = await resolveIssue(session.id, key, optionId);
      props.onReportUpdate({ ...data, report: { ...res.report } });
      if (
        (issue.ruleId === "APA-HEAD-002" || issue.ruleId === "APA-HEAD-003") &&
        optionId.startsWith("level")
      ) {
        setHeadingResolved(true);
      }
    } finally {
      setResolving(null);
    }
  };

  const regen = async () => {
    await regenerate(session.id);
    props.onRegenerate();
  };

  const validated = report.state === "apa_validated";
  const canDownload = session.status !== "error" && data.settings.mode !== "check";

  return (
    <div className="results-page">
      <div className="page-heading results-heading">
        <div><span className="eyebrow">Analysis complete</span><h1>Your APA 7 document report</h1></div>
        <p>{session.originalName}</p>
      </div>
      <div
        className={`state-banner ${validated ? "ok" : "review"}`}
        role="status"
        aria-live="polite"
      >
        <div>
          <h2>{validated ? "APA 7 Validated" : "Review Required"}</h2>
          <div className="sub">
            {validated
              ? "All applicable rules have been checked or resolved."
              : `${report.unresolvedCount} item${report.unresolvedCount === 1 ? "" : "s"} need your attention before validation can be completed.`}
          </div>
        </div>
        <div className="summary-stats">
          <div className="stat">
            <strong>{report.rulesResolvedPercent}%</strong>
            <span>rules resolved</span>
          </div>
          <div className="stat">
            <strong>{report.changesApplied}</strong>
            <span>formatting changes applied</span>
          </div>
          <div className="stat">
            <strong>{report.unresolvedCount}</strong>
            <span>unresolved issues</span>
          </div>
        </div>
      </div>

      {report.instructorOverrides.length > 0 && (
        <div className="card">
          <h2>Instructor overrides</h2>
          {report.instructorOverrides.map((o, i) => (
            <div key={i}>{o}</div>
          ))}
          {data.instructorUninterpreted.length > 0 && (
            <p className="section-note">
              Not automatically interpreted (please check manually):{" "}
              {data.instructorUninterpreted.join(" · ")}
            </p>
          )}
        </div>
      )}

      <div className="card">
        <h2>Category overview</h2>
        <div className="cat-grid">
          {report.categories.map((c) => (
            <div className="cat" key={c.category}>
              <div className="label">{c.label}</div>
              <span className={`badge ${c.status}`}>
                {STATUS_LABEL[c.status] ?? c.status}
              </span>{" "}
              <div className="meta">
                {c.rulesPassed}/{c.rulesChecked} rules · {c.issueCount} finding
                {c.issueCount === 1 ? "" : "s"}
              </div>
            </div>
          ))}
        </div>
      </div>

      {report.verification && (
        <div className="card">
          <h2>Reference verification</h2>
          <p>
            Provider: <strong>{report.verification.provider}</strong> —{" "}
            {report.verification.verified} verified,{" "}
            {report.verification.probable} probable,{" "}
            {report.verification.mismatched} mismatched,{" "}
            {report.verification.unverified} unverified.
          </p>
          {report.verification.providerUnavailable && (
            <p className="section-note">
              External metadata verification was temporarily unavailable for
              some entries — those are marked UNVERIFIED, not failed.
            </p>
          )}
        </div>
      )}

      <div className="card">
        <h2>Findings ({report.issues.length})</h2>
        <div className="filters" role="tablist" aria-label="Filter findings">
          {(
            [
              ["all", "All"],
              ["errors", "Errors"],
              ["warnings", "Warnings"],
              ["needs_review", "Needs review"],
            ] as [Filter, string][]
          ).map(([id, label]) => (
            <button
              key={id}
              className={`chip${filter === id ? " active" : ""}`}
              onClick={() => setFilter(id)}
              aria-pressed={filter === id}
            >
              {label}
            </button>
          ))}
        </div>

        {issues.length === 0 && <p>No findings for this filter.</p>}
        {issues.map((issue) => (
          <IssueCard
            key={issue.id}
            issue={issue}
            busy={resolving === issueKeyOf(issue)}
            onResolve={(optionId) => void resolve(issue, optionId)}
          />
        ))}
        {headingResolved && (
          <div className="section-note" style={{ marginTop: "0.6rem" }}>
            You confirmed heading levels.{" "}
            <button className="btn small primary" onClick={() => void regen()}>
              Re-apply formatting with your confirmations
            </button>
          </div>
        )}
      </div>

      <div className="card">
        <div className="collapsible-heading">
          <div>
            <h2>Applied changes ({report.changes.length})</h2>
            <p>Optional technical detail — hidden to keep your download within easy reach.</p>
          </div>
          <button className="btn small" aria-expanded={showChanges} onClick={() => setShowChanges((shown) => !shown)}>
            {showChanges ? "Hide changes" : `Show all ${report.changes.length} changes`}
          </button>
        </div>
        {report.changes.length === 0 && <p>No changes were applied.</p>}
        {showChanges && report.changes.map((c) => (
          <div className="change" key={c.id}>
            <span className="rule">{c.ruleId}</span>{" "}
            {c.location.description ??
              (c.location.paragraphIndex != null
                ? `Paragraph ${c.location.paragraphIndex + 1}`
                : "")}
            {c.location.excerpt ? ` — “${c.location.excerpt}”` : ""}
            <div className="diff">
              <span className="before">{c.before}</span>
              <span aria-hidden="true">→</span>
              <span className="after">{c.after}</span>
            </div>
            <div className="reason">{c.reason}</div>
          </div>
        ))}
        {showChanges && <p className="section-note">
          Want to undo a whole class of changes? Reprocess the pristine
          original — it is kept untouched for the retention period.
        </p>}
      </div>

      <div className="card">
        <h2>
          <button
            className="btn small"
            aria-expanded={showOutline}
            onClick={() => setShowOutline((s) => !s)}
          >
            {showOutline ? "Hide" : "Show"}
          </button>{" "}
          Document structure
        </h2>
        {showOutline && (
          <div className="outline">
            {data.outline.length === 0 && <p>No structural elements detected.</p>}
            {data.outline.map((e, i) => (
              <div className="entry" key={i}>
                <span className="kind">
                  {e.kind === "heading"
                    ? `Level ${e.level ?? "?"}`
                    : e.kind.replace(/-/g, " ")}
                  {e.confidence && e.confidence !== "high" ? ` (${e.confidence})` : ""}
                </span>
                <span className="txt">
                  {e.text}
                  {e.issues > 0 && (
                    <>
                      {" "}
                      <span className="badge warning">
                        {e.issues} finding{e.issues === 1 ? "" : "s"}
                      </span>
                    </>
                  )}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="card">
        <h2>Download</h2>
        <div className="dl-row">
          {canDownload ? (
            <a className="btn primary" href={downloadUrl(session.id)}>
              Download Corrected DOCX
            </a>
          ) : (
            <span className="section-note">
              Check mode does not modify the document, so there is no corrected
              file — run Format or Format + Verify to generate one.
            </span>
          )}
          <a className="btn" href={reportDownloadUrl(session.id)}>
            Download Compliance Report (HTML)
          </a>
          <button className="btn" onClick={props.onReset}>
            Start over with a new document
          </button>
        </div>
        {canDownload && (
          <p className="section-note">
            The corrected file downloads as{" "}
            <code>
              {session.originalName.replace(/\.docx$/i, "")}_
              {data.settings.mode === "format_verify"
                ? "APA7_verified"
                : "APA7_formatted"}
              .docx
            </code>
            . Your original upload remains unchanged.
          </p>
        )}
      </div>
    </div>
  );
}

function IssueCard(props: {
  issue: Issue;
  busy: boolean;
  onResolve: (optionId: string) => void;
}) {
  const { issue } = props;
  const needsAction = issue.userResolutionRequired && !issue.resolution;
  return (
    <div className={`issue ${issue.severity}`}>
      <div className="head">
        <span className={`badge ${issue.resolution ? "resolved" : issue.status}`}>
          {issue.resolution ? "RESOLVED" : STATUS_LABEL[issue.status] ?? issue.status}
        </span>
        <span className="msg">{issue.message}</span>
        <span className="rule">
          {issue.ruleId} · {Math.round(issue.confidence * 100)}%
        </span>
      </div>
      {issue.explanation && <p className="why">{issue.explanation}</p>}
      {issue.location?.excerpt && (
        <div className="where">
          near: “{issue.location.excerpt}”
          {issue.location.paragraphIndex != null
            ? ` (paragraph ${issue.location.paragraphIndex + 1})`
            : ""}
        </div>
      )}
      {issue.suggestedValue && !issue.resolution && (
        <div className="why">Suggested: {issue.suggestedValue}</div>
      )}
      {issue.resolution && (
        <div className="why">
          Your resolution:{" "}
          {issue.resolutionOptions?.find((o) => o.id === issue.resolution?.optionId)
            ?.label ?? issue.resolution.optionId}
        </div>
      )}
      {needsAction && issue.resolutionOptions && (
        <div className="resolve-row">
          {issue.resolutionOptions.map((o) => (
            <button
              key={o.id}
              className="btn small"
              disabled={props.busy}
              title={o.description}
              onClick={() => props.onResolve(o.id)}
            >
              {o.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
