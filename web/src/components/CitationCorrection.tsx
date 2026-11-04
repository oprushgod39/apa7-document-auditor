import { useRef, useState } from "react";

interface Result {
  text: string;
  citationsFound: number;
  changes: { before: string; after: string; reason: string }[];
  warnings: string[];
}
const EXAMPLE = "Research supports this finding (Smith and Patel 2024).\nKumar et al (2023) reported similar results.\nSmith & Patel (2024) explained the implications.";

export function CitationCorrectionScreen() {
  const [mode, setMode] = useState<"text" | "document">("text");
  const [text, setText] = useState("");
  const [result, setResult] = useState<Result | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const generation = useRef(0);
  const update = (value: string) => {
    generation.current++;
    setText(value); setResult(null); setError(""); setCopied(false);
  };
  const correct = async () => {
    const current = ++generation.current;
    setBusy(true); setError(""); setCopied(false);
    try {
      const response = await fetch("/api/citations/correct", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data?.error?.message ?? "Could not correct citations. Please try again.");
      if (generation.current === current) setResult(data as Result);
    } catch (e) {
      if (generation.current === current) setError(e instanceof Error ? e.message : "Could not correct citations.");
    } finally { setBusy(false); }
  };
  const copy = async () => {
    if (!result) return;
    try { await navigator.clipboard.writeText(result.text); setCopied(true); }
    catch { setError("Could not copy automatically. Select the corrected text and copy it manually."); }
  };
  return <div className="citation-workspace">
    <div className="page-heading">
      <div><span className="eyebrow">Citation workspace</span><h1>Correct your citations.</h1></div>
      <p>Paste text or upload a Word paper with its references. Correct APA 7 citations while preserving the surrounding wording.</p>
    </div>
    <div className="product-switch citation-mode" aria-label="Citation input method">
      <button className={mode === "text" ? "active" : ""} aria-pressed={mode === "text"} onClick={() => setMode("text")}>Paste text</button>
      <button className={mode === "document" ? "active" : ""} aria-pressed={mode === "document"} onClick={() => setMode("document")}>Upload Word document</button>
    </div>
    {mode === "document" ? <CitationDocument /> : <>
    <div className="citation-editors">
      <section className="card">
        <div className="citation-panel-title"><h2>Your text</h2><button className="btn small" disabled={busy} onClick={() => update(EXAMPLE)}>Try an example</button></div>
        <label htmlFor="citation-input">Paste text with in-text citations</label>
        <textarea id="citation-input" className="citation-textarea" value={text} maxLength={50_000}
          placeholder="For example: Research supports this finding (Smith and Patel 2024)."
          onChange={(e) => update(e.target.value)} aria-describedby="citation-input-help" />
        <div className="citation-panel-title"><small id="citation-input-help">{text.length.toLocaleString()} / 50,000 characters</small>
          <button className="btn small" disabled={busy || !text} onClick={() => update("")}>Clear</button></div>
        <button className="btn primary" disabled={busy || !text.trim()} onClick={correct}>{busy ? "Correcting…" : "Correct citations"}</button>
      </section>
      <section className="card" aria-busy={busy}>
        <div className="citation-panel-title"><h2>Corrected text</h2>
          <button className="btn small" disabled={!result || busy} onClick={copy}>{copied ? "Copied!" : "Copy text"}</button></div>
        <label htmlFor="citation-output">APA 7 citation formatting</label>
        <textarea id="citation-output" className="citation-textarea" readOnly value={result?.text ?? ""}
          placeholder="Your corrected text will appear here." />
        <p className="section-note" role="status">{result
          ? `${result.citationsFound} citation${result.citationsFound === 1 ? "" : "s"} detected · ${result.changes.length} correction${result.changes.length === 1 ? "" : "s"} applied`
          : "Author names, dates, and the surrounding wording are preserved."}</p>
      </section>
    </div>
    {error && <div className="error-box" role="alert">{error}</div>}
    {result && result.warnings.length > 0 && <section className="card"><h2>Check these items</h2>
      <ul>{result.warnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul></section>}
    {result && result.changes.length > 0 && <details className="card"><summary>View corrections ({result.changes.length})</summary>
      <ul className="citation-changes">{result.changes.map((change, i) => <li key={i}>
        <del>{change.before}</del><span aria-hidden="true"> → </span><strong>{change.after}</strong><p>{change.reason}</p>
      </li>)}</ul></details>}
    </>}
    <p className="section-note">Corrects citation mechanics such as author–year commas, “and” versus “&amp;”, “et al.” punctuation, and page locators. Supplied references help identify incomplete and ambiguous citations; source content is not verified. <a href="https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date" target="_blank" rel="noreferrer">APA citation guidance</a></p>
  </div>;
}

interface DocumentResult {
  id: string;
  downloadUrl: string;
  citationsFound: number;
  referencesFound: number;
  paragraphsCorrected: number;
  changes: { paragraph: number; before: string; after: string; reason: string }[];
  warnings: string[];
  uncitedReferences: string[];
}

function CitationDocument() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<DocumentResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState("");
  const select = (selected: File | null) => {
    setResult(null); setError(""); setFile(null);
    if (!selected) return;
    if (!/\.docx$/i.test(selected.name)) { setError("Choose a Word .docx document."); return; }
    if (selected.size > 4 * 1024 * 1024) { setError("Choose a DOCX file no larger than 4 MB."); return; }
    setFile(selected);
  };
  const correct = async () => {
    if (!file) return;
    setBusy(true); setError(""); setResult(null);
    try {
      const form = new FormData(); form.append("file", file);
      const response = await fetch("/api/citations/documents", { method: "POST", body: form });
      const data = await response.json();
      if (!response.ok) throw new Error(data?.error?.message ?? "Could not correct this document. Please try again.");
      setResult(data as DocumentResult);
    } catch (e) { setError(e instanceof Error ? e.message : "Document correction failed."); }
    finally { setBusy(false); }
  };
  const download = async () => {
    if (!result) return;
    setDownloading(true); setError("");
    try {
      const response = await fetch(result.downloadUrl);
      if (!response.ok) {
        const data = await response.json().catch(() => null);
        throw new Error(data?.error?.message ?? "Download failed. Please correct the document again.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a"); anchor.href = url;
      anchor.download = (file?.name.replace(/\.docx$/i, "") ?? "document") + "_citations_corrected.docx";
      anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 30_000);
    } catch (e) { setError(e instanceof Error ? e.message : "Could not download the document."); }
    finally { setDownloading(false); }
  };
  return <>
    <section className="card">
      <h2>Upload your paper with its reference list</h2>
      <p className="section-note">Citation corrections only. Your wording, document formatting, images, tables, and reference entries are preserved. Use a “References” heading before the reference list.</p>
      <label htmlFor="citation-document">Word document (.docx, up to 4 MB)</label>
      <input id="citation-document" className="citation-file" type="file" accept=".docx" disabled={busy || downloading} onChange={(e) => select(e.target.files?.[0] ?? null)} />
      <button className="btn primary" disabled={!file || busy || downloading} onClick={correct}>{busy ? "Correcting document…" : "Correct document citations"}</button>
      <p role="status" className="section-note">{busy ? "Checking citations against your references and preparing the corrected Word document…" : file?.name}</p>
    </section>
    {error && <div className="error-box" role="alert">{error}</div>}
    {result && <section className="card">
      <h2>Your corrected document is ready</h2>
      <p role="status">{result.citationsFound} citations detected · {result.referencesFound} references · {result.paragraphsCorrected} paragraphs corrected</p>
      <button className="btn primary" disabled={downloading} onClick={download}>{downloading ? "Downloading…" : "Download corrected Word document"}</button>
      <p className="section-note">Download now; files are automatically removed under the site's retention policy.</p>
      {result.warnings.length > 0 && <details open><summary>Items to review ({result.warnings.length})</summary><ul>{result.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul></details>}
      {result.uncitedReferences.length > 0 && <details><summary>References without matching citations ({result.uncitedReferences.length})</summary><ul>{result.uncitedReferences.map((r, i) => <li key={i}>{r}</li>)}</ul></details>}
      {result.changes.length > 0 && <details><summary>View corrections ({result.changes.length} paragraphs)</summary>
        <ul className="citation-changes">{result.changes.map((change, i) => <li key={i}><strong>Paragraph {change.paragraph}</strong><p><del>{change.before}</del></p><p>{change.after}</p></li>)}</ul>
      </details>}
    </section>}
  </>;
}
