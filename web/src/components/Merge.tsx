import { useEffect, useMemo, useRef, useState } from "react";
import { mergeDocuments, previewMergeDocuments } from "../lib/api";

interface MergeFile {
  id: string;
  file: File;
  name: string;
}

function suggestedName(filename: string): string {
  const stem = filename.replace(/\.docx$/i, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  return stem || "Student submission";
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
  const documentWords = useMemo(() => items.reduce((sum, item) => sum + (contentCounts[item.id] ?? 0) + countWords(item.name), 0), [items, contentCounts]);
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
        if (!known.has(key) && next.length < 30) {
          known.add(key);
          next.push({ id: crypto.randomUUID(), file, name: suggestedName(file.name) });
        }
      }
      if (current.length + docx.length > 30) setError("A merge can contain no more than 30 documents.");
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
    if (items.some((item) => !item.name.trim())) return setError("Each document needs a heading name.");
    setBusy(true);
    setError(null);
    try {
      const blob = await mergeDocuments(items.map(({ file, name }) => ({ file, name: name.trim() })), appendixWords);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "Merged_Submissions.docx";
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
          <p>Combine up to 30 Word documents in your chosen order while preserving their original formatting, tables, and images.</p>
        </div>
        <div className="merge-promise" aria-label="Merge rules">
          <span><strong>30</strong><small>documents max</small></span>
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
            <small>or click to choose files · 2–30 documents</small>
          </button>
          <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" multiple hidden onChange={(event) => addFiles([...(event.target.files ?? [])])} />

          {items.length > 0 && <div className="merge-queue">
            <div className="merge-queue-head"><div><span>{items.length}</span><strong>Documents in final order</strong></div><button onClick={() => setItems([])}>Clear all</button></div>
            {items.map((item, index) => <article className="merge-item" key={item.id}>
              <span className="merge-order">{String(index + 1).padStart(2, "0")}</span>
              <div className="merge-file-copy"><strong>{item.file.name}</strong><small>{(item.file.size / 1024 / 1024).toFixed(2)} MB · {counting ? "Counting…" : `${((contentCounts[item.id] ?? 0) + countWords(item.name)).toLocaleString()} words`} · References excluded</small></div>
              <label><small>Heading in merged file</small><input value={item.name} maxLength={200} onChange={(event) => setItems((current) => current.map((entry) => entry.id === item.id ? { ...entry, name: event.target.value } : entry))} /></label>
              <div className="merge-controls">
                <button aria-label={`Move ${item.file.name} up`} disabled={index === 0} onClick={() => move(index, -1)}>↑</button>
                <button aria-label={`Move ${item.file.name} down`} disabled={index === items.length - 1} onClick={() => move(index, 1)}>↓</button>
                <button className="remove" aria-label={`Remove ${item.file.name}`} onClick={() => setItems((current) => current.filter((entry) => entry.id !== item.id))}>×</button>
              </div>
            </article>)}
          </div>}
        </div>

        <aside className="merge-sidebar">
          <div className="merge-summary">
            <span className="eyebrow">Final document recipe</span>
            <ol>
              <li><span>1</span><div><strong>Centered document heading</strong><small>Editable above for every student or submission.</small></div></li>
              <li><span>2</span><div><strong>Original document content</strong><small>Formatting, tables, and images stay intact.</small></div></li>
              <li><span>3</span><div><strong>References removed</strong><small>References, Bibliography, or Works Cited through the end.</small></div></li>
              <li><span>4</span><div><strong>Controlled appendix</strong><small>The source repeats as needed and stops at your exact word target.</small></div></li>
            </ol>
            <div className="word-budget" aria-live="polite">
              <div><span>Before appendix</span><strong>{counting ? "Counting…" : documentWords.toLocaleString()}</strong><small>words · references excluded</small></div>
              <label><span>Appendix word target</span><input type="number" min={0} max={50000} step={100} value={appendixWords} onChange={(event) => setAppendixWords(Math.max(0, Math.min(50000, Number.parseInt(event.target.value || "0", 10))))} /></label>
              <div className="projected-total"><span>Projected final total</span><strong>{projectedWords.toLocaleString()}</strong><small>words · IGNORE label excluded</small></div>
              {appendixSourceWords > 0 && <p>Global-warming source: {appendixSourceWords.toLocaleString()} words. It repeats automatically when your target is longer.</p>}
            </div>
            <div className="merge-total"><span>Total upload</span><strong>{(totalSize / 1024 / 1024).toFixed(2)} MB</strong></div>
            {error && <p className="merge-error" role="alert">{error}</p>}
            <button className="btn primary merge-action" disabled={busy || counting || items.length < 2} onClick={runMerge}>{busy ? "Building your document…" : "Merge & download DOCX"}</button>
            <small className="merge-local-note">Save and close Microsoft Word first. A private Word process then performs the merge locally for maximum fidelity.</small>
          </div>
        </aside>
      </div>
    </section>
  );
}
