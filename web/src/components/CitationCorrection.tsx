import { useState } from "react";
import { CitationVerification } from "./CitationVerification";

export function CitationCorrectionScreen() {
  const [basic, setBasic] = useState(false);
  return basic ? <div className="citation-workspace"><button className="btn" onClick={() => setBasic(false)}>Back to claim verification</button><CitationDocument /></div>
    : <CitationVerification onBasic={() => setBasic(true)} />;
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
  uncitedParagraphs: { paragraph: number; excerpt: string }[];
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
      {result.uncitedParagraphs?.length > 0 && <details open><summary>Possible missing citations ({result.uncitedParagraphs.length})</summary><p className="section-note">These body paragraphs have no recognizable citation. Review the claims and choose a supporting reference before inserting one.</p><ol>{result.uncitedParagraphs.map((item) => <li key={item.paragraph}>Paragraph {item.paragraph}: {item.excerpt}{item.excerpt.length === 180 ? "…" : ""}</li>)}</ol></details>}
      {result.changes.length > 0 && <details><summary>View corrections ({result.changes.length} paragraphs)</summary>
        <ul className="citation-changes">{result.changes.map((change, i) => <li key={i}><strong>Paragraph {change.paragraph}</strong><p><del>{change.before}</del></p><p>{change.after}</p></li>)}</ul>
      </details>}
    </section>}
  </>;
}
