import { useEffect, useMemo, useRef, useState } from "react";
import { mergeDocuments, previewMergeDocuments } from "../lib/api";
import { parseRedactionTerms, repairPdf, type PdfRepairResult } from "../lib/pdfRedaction";

interface MergeFile {
  id: string;
  file: File;
}

const MERGE_MAX_DOCUMENTS = 50;

function anonymousHeading(index: number): string {
  return String(index + 1);
}

function countWords(text: string): number {
  return text.match(/[\p{L}\p{N}]+(?:['’.-][\p{L}\p{N}]+)*/gu)?.length ?? 0;
}

export function MergeScreen() {
  const [items, setItems] = useState<MergeFile[]>([]);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [counting, setCounting] = useState(false);
  const [contentCounts, setContentCounts] = useState<Record<string, number>>({});
  const [appendixSourceWords, setAppendixSourceWords] = useState(0);
  const [appendixWords, setAppendixWords] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const totalSize = useMemo(() => items.reduce((sum, item) => sum + item.file.size, 0), [items]);
  const fileSignature = useMemo(() => items.map((item) => `${item.id}:${item.file.size}:${item.file.lastModified}`).join("|"), [items]);
  const documentWords = useMemo(() => items.reduce((sum, item, index) => sum + (contentCounts[item.id] ?? 0) + countWords(anonymousHeading(index)), 0), [items, contentCounts]);
  const projectedWords = documentWords + appendixWords;

  useEffect(() => {
    if (items.length === 0) { setContentCounts({}); return; }
    const snapshot = [...items];
    const controller = new AbortController();
    setCounting(true);
    previewMergeDocuments(snapshot.map((item) => item.file))
      .then((preview) => {
        if (controller.signal.aborted) return;
        setContentCounts(Object.fromEntries(snapshot.map((item, index) => [item.id, preview.contentWords[index] ?? 0])));
        setAppendixSourceWords(preview.appendixSourceWords);
      })
      .catch((cause) => { if (!controller.signal.aborted) setError(cause instanceof Error ? cause.message : "Word count could not be calculated."); })
      .finally(() => { if (!controller.signal.aborted) setCounting(false); });
    return () => controller.abort();
  }, [fileSignature]);

  const addFiles = (selected: File[]) => {
    setError(null);
    const docx = selected.filter((file) => /\.docx$/i.test(file.name));
    if (docx.length !== selected.length) setError("Only Microsoft Word .docx files can preserve all formatting, tables, and images.");
    setItems((current) => {
      const known = new Set(current.map((item) => `${item.file.name}|${item.file.size}|${item.file.lastModified}`));
      const next = [...current];
      for (const file of docx) {
        const key = `${file.name}|${file.size}|${file.lastModified}`;
        if (!known.has(key) && next.length < MERGE_MAX_DOCUMENTS) {
          known.add(key);
          next.push({ id: crypto.randomUUID(), file });
        }
      }
      if (current.length + docx.length > MERGE_MAX_DOCUMENTS) setError(`A merge can contain no more than ${MERGE_MAX_DOCUMENTS} documents.`);
      return next;
    });
  };

  const move = (index: number, direction: -1 | 1) => {
    setItems((current) => {
      const target = index + direction;
      if (target < 0 || target >= current.length) return current;
      const next = [...current];
      [next[index], next[target]] = [next[target]!, next[index]!];
      return next;
    });
  };

  const runMerge = async () => {
    if (items.length < 2) return setError("Add at least two DOCX files to create a merged document.");
    setBusy(true);
    setError(null);
    try {
      const blob = await mergeDocuments(items.map(({ file }, index) => ({ file, name: anonymousHeading(index) })), appendixWords);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "Merged_Submissions.pdf";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 2_000);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The documents could not be merged.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className="merge-page">
      <div className="merge-hero">
        <div>
          <span className="eyebrow">Submission builder</span>
          <h1>One polished file.<br/><em>Every document intact.</em></h1>
          <p>Combine up to 50 Word documents in your chosen order while preserving their original formatting, tables, and images.</p>
        </div>
        <div className="merge-promise" aria-label="Merge rules">
          <span><strong>50</strong><small>documents max</small></span>
          <span><strong>0</strong><small>reference lists kept</small></span>
          <span><strong>{appendixWords.toLocaleString()}</strong><small>appendix words selected</small></span>
        </div>
      </div>

      <div className="merge-layout">
        <div className="merge-workspace">
          <button
            className={`merge-drop${dragging ? " dragging" : ""}`}
            onClick={() => inputRef.current?.click()}
            onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => { event.preventDefault(); setDragging(false); addFiles([...event.dataTransfer.files]); }}
          >
            <span className="merge-drop-icon">＋</span>
            <strong>Drop DOCX files here</strong>
            <small>or click to choose files · 2–50 documents</small>
          </button>
          <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden onChange={(event) => addFiles([...(event.target.files ?? [])])} />

          {items.length > 0 && <div className="merge-queue">
            <div className="merge-queue-head"><div><span>{items.length}</span><strong>Documents in final order</strong></div><button onClick={() => setItems([])}>Clear all</button></div>
            {items.map((item, index) => <article className="merge-item" key={item.id}>
              <span className="merge-order">{String(index + 1).padStart(2, "0")}</span>
              <div className="merge-file-copy"><strong>{anonymousHeading(index)}</strong><small>{(item.file.size / 1024 / 1024).toFixed(2)} MB · {counting ? "Counting…" : `${((contentCounts[item.id] ?? 0) + countWords(anonymousHeading(index))).toLocaleString()} words`} · Student details and references excluded</small></div>
              <label><small>Anonymous heading in merged file</small><input value={anonymousHeading(index)} readOnly aria-readonly="true" /></label>
              <div className="merge-controls">
                <button aria-label={`Move ${anonymousHeading(index)} up`} disabled={index === 0} onClick={() => move(index, -1)}>↑</button>
                <button aria-label={`Move ${anonymousHeading(index)} down`} disabled={index === items.length - 1} onClick={() => move(index, 1)}>↓</button>
                <button className="remove" aria-label={`Remove ${anonymousHeading(index)}`} onClick={() => setItems((current) => current.filter((entry) => entry.id !== item.id))}>×</button>
              </div>
            </article>)}
          </div>}
        </div>

        <aside className="merge-sidebar">
          <div className="merge-summary">
            <span className="eyebrow">Final document recipe</span>
            <ol>
              <li><span>1</span><div><strong>Numbered document heading</strong><small>Uses 1, 2, 3, and so on—never the uploaded filename.</small></div></li>
              <li><span>2</span><div><strong>Original document content</strong><small>Body formatting, tables, and images stay intact; title-page details are omitted.</small></div></li>
              <li><span>3</span><div><strong>References removed</strong><small>References, Bibliography, or Works Cited through the end.</small></div></li>
              <li><span>4</span><div><strong>Controlled appendix</strong><small>The source repeats as needed and stops at your exact word target.</small></div></li>
            </ol>
            <div className="word-budget" aria-live="polite">
              <div><span>Before appendix</span><strong>{counting ? "Counting…" : documentWords.toLocaleString()}</strong><small>words · student details and references excluded</small></div>
              <label><span>Appendix word target</span><input type="number" min={0} max={50000} step={100} value={appendixWords} onChange={(event) => setAppendixWords(Math.max(0, Math.min(50000, Number.parseInt(event.target.value || "0", 10))))} /></label>
              <div className="projected-total"><span>Projected final total</span><strong>{projectedWords.toLocaleString()}</strong><small>words · IGNORE label excluded</small></div>
              {appendixSourceWords > 0 && <p>Global-warming source: {appendixSourceWords.toLocaleString()} words. It repeats automatically when your target is longer.</p>}
            </div>
            <div className="merge-total"><span>Total upload</span><strong>{(totalSize / 1024 / 1024).toFixed(2)} MB</strong></div>
            {error && <p className="merge-error" role="alert">{error}</p>}
            <button className="btn primary merge-action" disabled={busy || counting || items.length < 2} onClick={runMerge}>{busy ? "Building your document…" : "Merge & download PDF"}</button>
            <small className="merge-local-note">No Microsoft Word needed — each document is rendered directly into one merged PDF, preserving formatting, tables, and images.</small>
          </div>
        </aside>
      </div>
      <PdfRepairPanel />
    </section>
  );
}

function PdfRepairPanel() {
  const [file, setFile] = useState<File | null>(null);
  const [terms, setTerms] = useState("");
  const [removeLabels, setRemoveLabels] = useState(true);
  const [replaceHeadings, setReplaceHeadings] = useState(true);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("Upload a faulty merged PDF to begin.");
  const [result, setResult] = useState<PdfRepairResult | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const choose = (selected: File | null) => {
    setResult(null);
    if (!selected) return setFile(null);
    if (selected.type !== "application/pdf" && !/\.pdf$/i.test(selected.name)) {
      setFile(null);
      setStatus("Choose a PDF file.");
      return;
    }
    setFile(selected);
    setStatus("Enter each student or university name exactly as it appears in the PDF.");
  };

  const run = async () => {
    if (!file) return;
    setBusy(true);
    setResult(null);
    try {
      const repaired = await repairPdf(file, parseRedactionTerms(terms), removeLabels, replaceHeadings, setStatus);
      setResult(repaired);
      const headingSummary = repaired.headingsReplaced > 0
        ? `${repaired.headingsReplaced} old filename heading${repaired.headingsReplaced === 1 ? "" : "s"}—including any names inside ${repaired.headingsReplaced === 1 ? "it" : "them"}—${repaired.headingsReplaced === 1 ? "was" : "were"} removed and replaced with numbers`
        : "No old filename headings were found";
      const matchSummary = repaired.matches > 0
        ? `${repaired.matches} additional name or title-page label match${repaired.matches === 1 ? "" : "es"} ${repaired.matches === 1 ? "was" : "were"} removed`
        : "No additional typed names or title-page labels were found";
      setStatus(`${headingSummary}. ${matchSummary}. ${repaired.words.toLocaleString()} searchable words preserved.`);
    } catch (cause) {
      setStatus(cause instanceof Error ? cause.message : "The PDF could not be repaired.");
    } finally {
      setBusy(false);
    }
  };

  const download = () => {
    if (!result || !file) return;
    const url = URL.createObjectURL(result.blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${file.name.replace(/\.pdf$/i, "")}_anonymized.pdf`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 2_000);
  };

  return <section className="pdf-repair">
    <div className="pdf-repair-copy">
      <span className="eyebrow">Optional PDF repair</span>
      <h2>Repair an older merged PDF.</h2>
      <p>Names are permanently removed from the PDF’s real text, while all remaining words stay searchable and countable. Old filename headings can also be replaced with 1, 2, 3, and so on. Processing stays in your browser.</p>
    </div>
    <div className="pdf-repair-form">
      <button className="pdf-repair-file" onClick={() => inputRef.current?.click()}>
        <span>PDF</span><div><strong>{file ? "PDF selected" : "Choose faulty PDF"}</strong><small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB` : "One PDF at a time"}</small></div>
      </button>
      <input ref={inputRef} type="file" accept=".pdf,application/pdf" hidden onChange={(event) => choose(event.target.files?.[0] ?? null)} />
      <label className="pdf-repair-terms"><span>Additional names and universities to remove</span><textarea rows={4} value={terms} onChange={(event) => { setTerms(event.target.value); setResult(null); }} placeholder={"Mahesh\nRiverside University\nAnother student name"} /><small>Names inside old filename headings are removed automatically. Enter a name here only if it also appears elsewhere in the PDF. Use one exact name per line; commas and semicolons also work.</small></label>
      <label className="pdf-repair-labels"><input type="checkbox" checked={removeLabels} onChange={(event) => { setRemoveLabels(event.target.checked); setResult(null); }} /><span><strong>Also remove default title-page labels</strong><small>Name, University, Subject code and name, Professor, and Due date.</small></span></label>
      <label className="pdf-repair-labels"><input type="checkbox" checked={replaceHeadings} onChange={(event) => { setReplaceHeadings(event.target.checked); setResult(null); }} /><span><strong>Replace old document headings with numbers</strong><small>The entire old filename heading—including any student name—is permanently removed, then replaced with 1, 2, 3, and so on.</small></span></label>
      <div className="pdf-repair-actions">
        <button className="btn primary" disabled={!file || busy || (parseRedactionTerms(terms).length === 0 && !removeLabels && !replaceHeadings)} onClick={() => void run()}>{busy ? "Repairing PDF…" : "Repair PDF"}</button>
        <button className="btn" disabled={!result || busy} onClick={download}>Download anonymized PDF</button>
      </div>
      <p className={`pdf-repair-status${busy ? " busy" : ""}`} aria-live="polite"><span />{status}</p>
    </div>
  </section>;
}
