import type { SessionInfo, UploadResponse } from "../lib/api";

const ALL_STAGES: [string, string][] = [
  ["read", "Reading Word document"],
  ["structure", "Analyzing document structure"],
  ["headings", "Detecting headings"],
  ["page_format", "Checking page formatting"],
  ["citations", "Analyzing citations"],
  ["references", "Analyzing references"],
  ["apply", "Applying APA corrections"],
  ["prepare_output", "Preparing corrected Word document"],
  ["verify_metadata", "Verifying scholarly metadata"],
  ["audit", "Running independent APA audit"],
];

export function ProcessingScreen(props: {
  session: UploadResponse;
  status: SessionInfo | null;
}) {
  const stages = props.status?.stages ?? [];
  const stateOf = (key: string) =>
    stages.find((s) => s.key === key)?.status ?? "pending";

  return (
    <div className="processing-wrap">
      <div className="processing-visual" aria-hidden="true"><span>A7</span><i /><i /><i /></div>
      <div className="card processing-card">
      <span className="eyebrow">Precision check in progress</span>
      <h2>Analyzing your document</h2>
      <p className="processing-file">{props.session.originalName}</p>
      <ul className="stages" aria-live="polite">
        {ALL_STAGES.map(([key, label]) => {
          const st = stateOf(key);
          return (
            <li key={key} className={st}>
              <span className="stage-dot" aria-hidden="true" />
              {label}
              <span className="sr-only">
                {st === "done" ? " — complete" : st === "running" ? " — in progress" : ""}
              </span>
            </li>
          );
        })}
      </ul>
      <p className="section-note">
        These stages reflect actual processing — no simulated progress.
      </p>
      </div>
    </div>
  );
}
