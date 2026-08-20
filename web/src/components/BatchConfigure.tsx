import { useState } from "react";
import type { ProcessSettings } from "../lib/api";
import type { BatchFileEntry } from "./BatchUpload";

const MODES = [
  {
    id: "format_verify",
    name: "Format + Verify",
    desc: "Apply safe APA corrections, audit citations & references, verify scholarly metadata. Recommended.",
  },
  {
    id: "format",
    name: "Format APA 7",
    desc: "Apply safe APA formatting rules and produce a corrected document.",
  },
  {
    id: "check",
    name: "Check APA 7",
    desc: "Analyze and report only — the document is not modified.",
  },
] as const;

export function BatchConfigureScreen(props: {
  files: BatchFileEntry[];
  onStart: (settings: ProcessSettings) => void;
  onBack: () => void;
}) {
  const [paperType, setPaperType] = useState<"student" | "professional">("student");
  const [mode, setMode] = useState<"check" | "format" | "format_verify">("format_verify");
  const [preserveWording] = useState(true);
  const [fixCitationMechanics, setFixCitationMechanics] = useState(true);
  const [verifyMetadata, setVerifyMetadata] = useState(true);
  const [showOptional, setShowOptional] = useState(false);
  const [meta, setMeta] = useState<Record<string, string>>({
    institution: "University",
    courseNumber: "",
    courseName: "subject code and name",
    instructor: "Professor _____",
    dueDate: "Due date",
    runningHead: "",
  });
  const [instructorRequirements, setInstructorRequirements] = useState("");

  const setM = (k: string, v: string) => setMeta((m) => ({ ...m, [k]: v }));

  const start = () => {
    const metadata: Record<string, string> = {};
    for (const [k, v] of Object.entries(meta)) {
      if (v.trim()) metadata[k] = v.trim();
    }
    props.onStart({
      paperType,
      mode,
      preserveWording,
      fixCitationMechanics,
      verifyMetadata,
      metadata,
      instructorRequirements,
    });
  };

  return (
    <div className="configure-page batch-configure-page">
      <div className="page-heading">
        <div>
          <span className="eyebrow">Batch workspace</span>
          <h1>Apply one set of settings to all {props.files.length} documents</h1>
        </div>
        <p>
          Each file keeps its own detected title, author, and structure. A title
          page is only ever created from a title actually detected in that
          document — nothing is invented. Shared assignment details below (course,
          instructor, due date) apply to every file in the batch.
        </p>
      </div>

      <div className="card">
        <h2>Documents ({props.files.length})</h2>
        <button className="btn small" onClick={props.onBack}>
          Change file selection
        </button>
        <ul className="batch-file-list">
          {props.files.map((item) => (
            <li key={item.id}>{item.file.name}</li>
          ))}
        </ul>
      </div>

      <div className="card">
        <h2>Paper type</h2>
        <div className="radio-row" role="radiogroup" aria-label="Paper type">
          {(
            [
              ["student", "Student Paper", "Title page with course, instructor and due date. No running head by default."],
              ["professional", "Professional Paper", "Running head, author note and abstract support."],
            ] as const
          ).map(([id, name, desc]) => (
            <label key={id} className={`radio-card${paperType === id ? " selected" : ""}`}>
              <input
                type="radio"
                name="batchPaperType"
                value={id}
                checked={paperType === id}
                onChange={() => setPaperType(id)}
              />
              <span>
                <strong>{name}</strong>
                <span>{desc}</span>
              </span>
            </label>
          ))}
        </div>
      </div>

      <div className="card">
        <h2>Mode</h2>
        <div className="radio-row" role="radiogroup" aria-label="Processing mode">
          {MODES.map((m) => (
            <label key={m.id} className={`radio-card${mode === m.id ? " selected" : ""}`}>
              <input
                type="radio"
                name="batchMode"
                value={m.id}
                checked={mode === m.id}
                onChange={() => setMode(m.id)}
              />
              <span>
                <strong>{m.name}</strong>
                <span>{m.desc}</span>
              </span>
            </label>
          ))}
        </div>
        <div style={{ marginTop: "1rem" }}>
          <div className="switch-row">
            <input type="checkbox" id="batch-preserve" checked readOnly disabled />
            <label htmlFor="batch-preserve">
              <strong>Preserve wording — ON</strong>
              <div className="desc">
                Your academic writing is never rewritten, paraphrased, or
                shortened. Only formatting and citation/reference mechanics are
                touched.
              </div>
            </label>
          </div>
          <div className="switch-row">
            <input
              type="checkbox"
              id="batch-citmech"
              checked={fixCitationMechanics}
              onChange={(e) => setFixCitationMechanics(e.target.checked)}
            />
            <label htmlFor="batch-citmech">
              Correct citation &amp; reference mechanics
              <div className="desc">
                Deterministic fixes such as “and” → “&amp;” in parentheses,
                “et al” → “et al.”, and modern DOI URLs.
              </div>
            </label>
          </div>
          {mode === "format_verify" && (
            <div className="switch-row">
              <input
                type="checkbox"
                id="batch-verifymeta"
                checked={verifyMetadata}
                onChange={(e) => setVerifyMetadata(e.target.checked)}
              />
              <label htmlFor="batch-verifymeta">
                Verify references against scholarly metadata (Crossref)
                <div className="desc">
                  Reference titles/authors/years are checked against public
                  scholarly records. Nothing is replaced without your
                  confirmation.
                </div>
              </label>
            </div>
          )}
        </div>
      </div>

      <div className="card">
        <h2>
          <button
            className="btn small"
            aria-expanded={showOptional}
            onClick={() => setShowOptional((s) => !s)}
          >
            {showOptional ? "Hide" : "Show"} optional settings
          </button>{" "}
          Shared title-page details
        </h2>
        {showOptional && (
          <>
            <p className="section-note">
              Applied to every document's title page (when one is created). Paper
              title and author always come from each document itself — never from
              here — so a file with no detectable title simply won't get an
              auto-generated title page; it's flagged in that file's row so you can
              format it individually if needed.
            </p>
            <div className="grid-2" style={{ marginTop: "0.8rem" }}>
              <div className="field">
                <label htmlFor="bm-inst">Institution</label>
                <input id="bm-inst" value={meta.institution} onChange={(e) => setM("institution", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="bm-cnum">Course number</label>
                <input id="bm-cnum" value={meta.courseNumber} onChange={(e) => setM("courseNumber", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="bm-cname">Course name</label>
                <input id="bm-cname" value={meta.courseName} onChange={(e) => setM("courseName", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="bm-instr">Instructor</label>
                <input id="bm-instr" value={meta.instructor} onChange={(e) => setM("instructor", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="bm-date">Due date</label>
                <input id="bm-date" value={meta.dueDate} onChange={(e) => setM("dueDate", e.target.value)} />
              </div>
              {paperType === "professional" && (
                <div className="field">
                  <label htmlFor="bm-rh">Running head (max 50 characters)</label>
                  <input
                    id="bm-rh"
                    maxLength={50}
                    value={meta.runningHead}
                    onChange={(e) => setM("runningHead", e.target.value)}
                  />
                </div>
              )}
            </div>
          </>
        )}
      </div>

      <div className="card">
        <h2>Instructor / assignment requirements (optional)</h2>
        <p className="section-note">
          Applied to every document in this batch. These are layered on top of
          the APA 7 baseline — deviations are labeled “Instructor override” in
          each document's report.
        </p>
        <div className="field" style={{ marginTop: "0.8rem" }}>
          <textarea
            id="batch-req"
            rows={4}
            placeholder={"e.g.\nTimes New Roman 12 only\nAbstract required\nMinimum five references"}
            value={instructorRequirements}
            onChange={(e) => setInstructorRequirements(e.target.value)}
          />
        </div>
      </div>

      <div className="action-dock">
        <div>
          <strong>Ready to format {props.files.length} documents?</strong>
          <span>Each original file stays unchanged.</span>
        </div>
        <button className="btn primary btn-large" onClick={start}>
          Start batch processing <span aria-hidden="true">→</span>
        </button>
      </div>
    </div>
  );
}
