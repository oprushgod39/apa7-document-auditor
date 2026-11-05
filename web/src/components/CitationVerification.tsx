import { useEffect, useRef, useState } from "react";

type Progress = { phase: "sources" | "claims" | "finalize" | "complete"; version: number; sourcesTotal: number; sourcesDone: number; paragraphsTotal: number; paragraphsDone: number; applied: number; review: number };
type Job = { id: string; token: string };
const KEY = "scholarlyworks.citation.verification";
export function CitationVerification({ onBasic }: { onBasic: () => void }) {
  const [file, setFile] = useState<File | null>(null), [job, setJob] = useState<Job | null>(null);
  const [progress, setProgress] = useState<Progress | null>(null), [busy, setBusy] = useState(false);
  const [available, setAvailable] = useState<boolean | null>(null), [error, setError] = useState("");
  const alive = useRef(true), running = useRef(false);
  useEffect(() => {
    alive.current = true;
    void fetch("/api/citations/verification/availability").then((r) => r.json()).then((r) => { if (alive.current) setAvailable(r.available === true); }).catch(() => { if (alive.current) setAvailable(false); });
    try { const saved = JSON.parse(sessionStorage.getItem(KEY) ?? "null"); if (saved?.id && saved?.token) setJob(saved); } catch { /* unavailable browser storage */ }
    return () => { alive.current = false; };
  }, []);
  const call = async (path: string, active: Job, body?: object) => {
    const response = await fetch(`/api/citations/verification/${active.id}${path}`, { method: body ? "POST" : "GET",
      headers: { "X-Citation-Token": active.token, ...(body ? { "Content-Type": "application/json" } : {}) }, body: body ? JSON.stringify(body) : undefined });
    const data = await response.json();
    if (!response.ok) throw new Error(data?.error?.message ?? "Verification could not continue. Retry to resume.");
    return data as Progress;
  };
  const resume = async (active: Job) => {
    if (running.current) return;
    running.current = true; setBusy(true); setError("");
    try {
      let state = await call("", active);
      while (alive.current) {
        setProgress(state);
        if (state.phase === "complete") break;
        state = await call("/step", active, { version: state.version });
      }
    } catch (e) { if (alive.current) setError(e instanceof Error ? e.message : "Verification failed. Retry to resume."); }
    finally { running.current = false; if (alive.current) setBusy(false); }
  };
  const upload = async () => {
    if (!file) return;
    setBusy(true); setError(""); setProgress(null);
    try {
      const form = new FormData(); form.append("file", file);
      const response = await fetch("/api/citations/verification", { method: "POST", body: form });
      const data = await response.json();
      if (!response.ok) throw new Error(data?.error?.message ?? "Could not read the document.");
      const active = { id: data.id, token: data.token }; setJob(active); setProgress(data);
      try { sessionStorage.setItem(KEY, JSON.stringify(active)); } catch { /* still usable this session */ }
      await resume(active);
    } catch (e) { setError(e instanceof Error ? e.message : "Upload failed."); }
    finally { if (alive.current) setBusy(false); }
  };
  const download = async (kind: "document" | "report") => {
    if (!job) return;
    try {
      const res = await fetch(`/api/citations/verification/${job.id}/${kind}`, { headers: { "X-Citation-Token": job.token } });
      if (!res.ok) throw new Error("Download unavailable. Your saved session may have expired.");
      const url = URL.createObjectURL(await res.blob()), a = document.createElement("a"); a.href = url;
      a.download = kind === "document" ? "citations_verified.docx" : "Citation_Verification_Report.html"; a.click(); setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (e) { setError(e instanceof Error ? e.message : "Download failed."); }
  };
  return <div className="citation-workspace"><div className="page-heading"><div><span className="eyebrow">Citation workspace</span><h1>Check claims. Correct citations.</h1></div><p>Upload your Word paper and references. We examine available source evidence, apply supported citation corrections, and prepare a review report.</p></div>
    <section className="card"><h2>Upload your document</h2><p className="section-note">Document passages and retrieved source text are sent to OpenAI for claim analysis. Sources that cannot be accessed, uncertain matches, and quotation locators are reported for review.</p>
      {available === false && <p role="status">Source verification is not configured yet. The site owner must add the OpenAI API key and citation model on the server.</p>}
      <p className="section-note"><button className="btn small" disabled={busy} onClick={onBasic}>Correct existing citation formatting only</button> This option works without source verification.</p>
      <label htmlFor="citation-upload">Word document with its reference list (.docx, up to 4 MB)</label>
      <input id="citation-upload" className="citation-file" type="file" accept=".docx" disabled={busy} onChange={(e) => { const f = e.target.files?.[0]; setError(""); if (f && (!/\.docx$/i.test(f.name) || f.size > 4 * 1024 * 1024)) { setError("Choose a DOCX no larger than 4 MB."); setFile(null); } else setFile(f ?? null); }} />
      <button className="btn primary" disabled={!file || busy || available !== true} onClick={() => void upload()}>{busy ? "Checking sources and claims…" : "Verify and correct citations"}</button>
      {job && !busy && progress?.phase !== "complete" && <button className="btn" onClick={() => void resume(job)}>Resume previous document</button>}
      {progress && <p role="status">{progress.phase === "sources" ? `Reading sources: ${progress.sourcesDone} of ${progress.sourcesTotal}` : progress.phase === "claims" ? `Checking document passages: ${progress.paragraphsDone} of ${progress.paragraphsTotal}` : progress.phase === "finalize" ? "Applying citations and checking document integrity…" : `Ready: ${progress.applied} evidence-backed citation edits · ${progress.review} claims to review`}</p>}
      {error && <p className="error-box" role="alert">{error}</p>}
    </section>
    {progress?.phase === "complete" && <section className="card"><h2>Your files are ready</h2><button className="btn primary" onClick={() => void download("document")}>Download corrected Word document</button> <button className="btn" onClick={() => void download("report")}>Download verification report</button><p className="section-note">The report contains claim decisions, evidence excerpts, source access limits, and unresolved items. Download both files before the session expires.</p></section>}
  </div>;
}
