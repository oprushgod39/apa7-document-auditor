import { useState } from "react";
import type { ProcessSettings, UploadResponse } from "../lib/api";

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

export function ConfigureScreen(props: {
  session: UploadResponse;
  onStart: (settings: ProcessSettings) => void;
  onReset: () => void;
}) {
  const d = props.session.detected;
  const [paperType, setPaperType] = useState<"student" | "professional">("student");
  const [mode, setMode] = useState<"check" | "format" | "format_verify">("format_verify");
  const [preserveWording] = useState(true);
  const [fixCitationMechanics, setFixCitationMechanics] = useState(true);
  const [verifyMetadata, setVerifyMetadata] = useState(true);
  const [showOptional, setShowOptional] = useState(false);
  const [meta, setMeta] = useState<Record<string, string>>({
    title: d.metadata.title ?? "",
    author: d.metadata.author ?? "",
    institution: d.metadata.institution ?? "",
    courseNumber: d.metadata.courseNumber ?? "",
    courseName: d.metadata.courseName ?? "",
    instructor: d.metadata.instructor ?? "",
    dueDate: d.metadata.dueDate ?? "",
    runningHead: "",
  });
  const [instructorRequirements, setInstructorRequirements] = useState("");

  const setM = (k: string, v: string) => setMeta((m) => ({ ...m, [k]: v }));
  const detected = (k: keyof typeof d.metadata) =>
    d.metadata[k] ? <span className="detected">detected in document</span> : null;

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
    <div>
      <div className="card">
        <h2>Document</h2>
        <p>
          <strong>{props.session.originalName}</strong>{" "}
          <button className="btn small" onClick={props.onReset} style={{ marginLeft: "0.6rem" }}>
            Choose a different file
          </button>
        </p>
        <div className="summary-stats" aria-label="Detected document structure">
          <div className="stat"><strong>{d.paragraphs}</strong><span>paragraphs</span></div>
          <div className="stat"><strong>{d.citations}</strong><span>citations</span></div>
          <div className="stat"><strong>{d.references}</strong><span>references</span></div>
          <div className="stat"><strong>{d.headings}</strong><span>headings</span></div>
          <div className="stat"><strong>{d.tables}</strong><span>tables</span></div>
          <div className="stat"><strong>{d.images}</strong><span>images</span></div>
        </div>
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
                name="paperType"
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
                name="mode"
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
            <input type="checkbox" id="preserve" checked readOnly disabled />
            <label htmlFor="preserve">
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
              id="citmech"
              checked={fixCitationMechanics}
              onChange={(e) => setFixCitationMechanics(e.target.checked)}
            />
            <label htmlFor="citmech">
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
                id="verifymeta"
                checked={verifyMetadata}
                onChange={(e) => setVerifyMetadata(e.target.checked)}
              />
              <label htmlFor="verifymeta">
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
          Title-page metadata &amp; assignment requirements
        </h2>
        {showOptional && (
          <>
            <p className="section-note">
              Detected values are pre-filled. Anything you enter here is used to
              complete a missing title page — the auditor never invents
              information.
            </p>
            <div className="grid-2" style={{ marginTop: "0.8rem" }}>
              <div className="field">
                <label htmlFor="m-title">Paper title {detected("title")}</label>
                <input id="m-title" value={meta.title} onChange={(e) => setM("title", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-author">Author name {detected("author")}</label>
                <input id="m-author" value={meta.author} onChange={(e) => setM("author", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-inst">Institution {detected("institution")}</label>
                <input id="m-inst" value={meta.institution} onChange={(e) => setM("institution", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-cnum">Course number {detected("courseNumber")}</label>
                <input id="m-cnum" value={meta.courseNumber} onChange={(e) => setM("courseNumber", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-cname">Course name {detected("courseName")}</label>
                <input id="m-cname" value={meta.courseName} onChange={(e) => setM("courseName", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-instr">Instructor {detected("instructor")}</label>
                <input id="m-instr" value={meta.instructor} onChange={(e) => setM("instructor", e.target.value)} />
              </div>
              <div className="field">
                <label htmlFor="m-date">Due date {detected("dueDate")}</label>
                <input id="m-date" value={meta.dueDate} onChange={(e) => setM("dueDate", e.target.value)} />
              </div>
              {paperType === "professional" && (
                <div className="field">
                  <label htmlFor="m-rh">Running head (max 50 characters)</label>
                  <input
                    id="m-rh"
                    maxLength={50}
                    value={meta.runningHead}
                    onChange={(e) => setM("runningHead", e.target.value)}
                  />
                </div>
              )}
            </div>
            <div className="field" style={{ marginTop: "1rem" }}>
              <label htmlFor="m-req">Instructor / assignment requirements (optional)</label>
              <textarea
                id="m-req"
                rows={4}
                placeholder={"e.g.\nTimes New Roman 12 only\nAbstract required\nMinimum five references"}
                value={instructorRequirements}
                onChange={(e) => setInstructorRequirements(e.target.value)}
              />
              <div className="desc section-note">
                These are layered on top of the APA 7 baseline. Deviations are
                labeled “Instructor override” in the report.
              </div>
            </div>
          </>
        )}
      </div>

      <button className="btn primary" onClick={start}>
        Start APA Analysis
      </button>
    </div>
  );
}
