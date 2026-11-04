import { useRef, useState, type ClipboardEvent } from "react";
import { sanitizeCitationHtml, citationHtmlText, plainCitationHtml, correctedCitationHtml, applyReferenceItalics } from "../lib/rich_citations";

interface Result {
  text: string;
  html: string;
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
  const [restoring, setRestoring] = useState(false);
  const [formatNotice, setFormatNotice] = useState("");
  const generation = useRef(0);
  const inputRef = useRef<HTMLDivElement | null>(null);
  const sourceHtml = useRef("");
  const [formatting, setFormatting] = useState({ italic: false, bold: false, underline: false });
  const syncFormatting = () => setFormatting({
    italic: document.queryCommandState("italic"), bold: document.queryCommandState("bold"), underline: document.queryCommandState("underline"),
  });
  const format = (command: "italic" | "bold" | "underline") => {
    inputRef.current?.focus();
    document.execCommand(command, false);
    captureInput(); syncFormatting();
  };
  const update = (value: string) => {
    setFormatNotice("");
    sourceHtml.current = plainCitationHtml(value);
    if (inputRef.current) inputRef.current.innerHTML = sourceHtml.current;
    generation.current++;
    setText(value); setResult(null); setError(""); setCopied(false);
  };
  const captureInput = () => {
    setFormatNotice("");
    sourceHtml.current = sanitizeCitationHtml(inputRef.current?.innerHTML ?? "");
    generation.current++;
    setText(citationHtmlText(sourceHtml.current)); setResult(null); setError(""); setCopied(false);
  };
  const restoreItalics = async () => {
    const original = sourceHtml.current;
    const value = citationHtmlText(original);
    if (!value.trim() || value.length > 50_000) return;
    const current = ++generation.current;
    setRestoring(true); setError("");
    try {
      const response = await fetch("/api/citations/reference-italics", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: value }) });
      const data = await response.json();
      if (!response.ok) throw new Error(data?.error?.message ?? "Could not restore reference italics.");
      if (current !== generation.current) return;
      const formatted = applyReferenceItalics(original, data.spans);
      sourceHtml.current = formatted;
      if (inputRef.current) inputRef.current.innerHTML = formatted;
      setResult(null); setCopied(false);
      setFormatNotice(data.spans.length ? "APA italics restored for recognized reference titles, journals, and volumes. Please review them." : data.warnings.join(" "));
    } catch (e) { if (current === generation.current) setError(e instanceof Error ? e.message : "Could not restore italics."); }
    finally { setRestoring(false); }
  };
  const paste = (event: ClipboardEvent<HTMLDivElement>) => {
    event.preventDefault();
    const html = event.clipboardData.getData("text/html");
    const safe = html ? sanitizeCitationHtml(html) : plainCitationHtml(event.clipboardData.getData("text/plain"));
    const selection = window.getSelection();
    const root = inputRef.current;
    if (!root || !selection?.rangeCount) return;
    const range = selection.getRangeAt(0);
    if (!root.contains(range.commonAncestorContainer)) return;
    range.deleteContents();
    const template = document.createElement("template"); template.innerHTML = safe;
    const last = template.content.lastChild;
    range.insertNode(template.content);
    if (last) { range.setStartAfter(last); range.collapse(true); selection.removeAllRanges(); selection.addRange(range); }
    captureInput();
    if (!/<(?:em|i)\b/i.test(sourceHtml.current) && /^\s*References\s*$/im.test(citationHtmlText(sourceHtml.current))) void restoreItalics();
  };
  const correct = async () => {
    if (text.length > 50_000) { setError("Use up to 50,000 characters, or upload your Word document."); return; }
    const current = ++generation.current;
    const originalHtml = sourceHtml.current;
    setBusy(true); setError(""); setCopied(false);
    try {
      const response = await fetch("/api/citations/correct", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data?.error?.message ?? "Could not correct citations. Please try again.");
      const formattingResponse = await fetch("/api/citations/reference-italics", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: data.text }) });
      if (!formattingResponse.ok) throw new Error("Citation text was corrected, but reference formatting could not be checked. Please try again.");
      const referenceFormatting = await formattingResponse.json();
      if (generation.current === current) setResult({ ...data, html: applyReferenceItalics(correctedCitationHtml(originalHtml, data.text), referenceFormatting.spans) } as Result);
    } catch (e) {
      if (generation.current === current) setError(e instanceof Error ? e.message : "Could not correct citations.");
    } finally { setBusy(false); }
  };
  const copy = async () => {
    if (!result) return;
    try {
      if (!navigator.clipboard.write || typeof ClipboardItem === "undefined") throw new Error("Formatted clipboard unavailable");
      await navigator.clipboard.write([new ClipboardItem({
        "text/plain": new Blob([result.text], { type: "text/plain" }),
        "text/html": new Blob([result.html], { type: "text/html" }),
      })]);
      setCopied(true);
    }
    catch { setError("Formatted copy is unavailable in this browser. Select the corrected text and copy it manually to retain italics, or use Word upload."); }
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
        <label id="citation-input-label" htmlFor="citation-input">Paste text with in-text citations and references</label>
        <div className="citation-format-toolbar" role="group" aria-label="Text formatting">
          <button type="button" className="btn small" aria-label="Bold" aria-pressed={formatting.bold} title="Bold (Ctrl+B)" onMouseDown={(e) => e.preventDefault()} onClick={() => format("bold")}><strong>B</strong></button>
          <button type="button" className="btn small" aria-label="Italic" aria-pressed={formatting.italic} title="Italic (Ctrl+I)" onMouseDown={(e) => e.preventDefault()} onClick={() => format("italic")}><em>I</em><span>Italic</span></button>
          <button type="button" className="btn small" aria-label="Underline" aria-pressed={formatting.underline} title="Underline (Ctrl+U)" onMouseDown={(e) => e.preventDefault()} onClick={() => format("underline")}><u>U</u></button>
          <button type="button" className="btn small" disabled={!text.trim() || busy || restoring} onClick={() => void restoreItalics()}>{restoring ? "Restoring italics…" : "Restore APA reference italics"}</button>
          <small>Select text to format it, or turn Italic on before typing.</small>
        </div>
        {formatNotice && <p className="section-note" role="status">{formatNotice}</p>}
        <div id="citation-input" className="citation-textarea citation-rich" contentEditable suppressContentEditableWarning
          role="textbox" aria-multiline="true" aria-labelledby="citation-input-label" aria-describedby="citation-input-help"
          data-placeholder="Paste from Word here. Italics and bold will be retained."
          ref={(node) => { inputRef.current = node; if (node && !node.innerHTML && sourceHtml.current) node.innerHTML = sourceHtml.current; }}
          onInput={() => { captureInput(); syncFormatting(); }} onKeyUp={syncFormatting} onMouseUp={syncFormatting} onFocus={syncFormatting} onPaste={paste} onDrop={(event) => event.preventDefault()} />
        <div className="citation-panel-title"><small id="citation-input-help">{text.length.toLocaleString()} / 50,000 characters</small>
          <button className="btn small" disabled={busy || !text} onClick={() => update("")}>Clear</button></div>
        <p className="section-note">If the source copies as plain text, use “Restore APA reference italics.” Recognizable reference titles, journal names, and volumes are also formatted in the corrected output.</p>
        <button className="btn primary" disabled={busy || restoring || !text.trim() || text.length > 50_000} onClick={correct}>{busy ? "Correcting…" : "Correct citations"}</button>
      </section>
      <section className="card" aria-busy={busy}>
        <div className="citation-panel-title"><h2>Corrected text</h2>
          <button className="btn small" disabled={!result || busy} onClick={copy}>{copied ? "Copied with formatting!" : "Copy with formatting"}</button></div>
        <div id="citation-output" className="citation-textarea citation-rich" role="region" aria-label="Corrected text with formatting" tabIndex={0}
          data-placeholder="Your corrected text will appear here."
          dangerouslySetInnerHTML={{ __html: result?.html ?? "" }} />
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
