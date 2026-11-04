import { useRef, useState } from "react";

interface Result {
  text: string;
  citationsFound: number;
  changes: { before: string; after: string; reason: string }[];
  warnings: string[];
}
const EXAMPLE = "Research supports this finding (Smith and Patel 2024).\nKumar et al (2023) reported similar results.\nSmith & Patel (2024) explained the implications.";

export function CitationCorrectionScreen() {
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
      <p>Paste a sentence, paragraph, or paper. Correct recognizable APA 7 in-text citation formatting, then copy the result.</p>
    </div>
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
    <p className="section-note">Corrects citation mechanics such as author–year commas, “and” versus “&amp;”, and “et al.” punctuation. It does not verify sources or add missing author names or years. <a href="https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date" target="_blank" rel="noreferrer">APA citation guidance</a></p>
  </div>;
}
