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
    <div className="card" style={{ maxWidth: "32rem", margin: "2rem auto" }}>
      <h2>Analyzing {props.session.originalName}</h2>
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
  );
}
