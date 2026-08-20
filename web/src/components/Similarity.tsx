import { useMemo, useRef, useState, type CSSProperties } from "react";
import {
  band,
  compareFiles,
  comparePassages,
  percent,
  type PassageMatch,
  type SimilarityDocument,
  type SimilarityResult,
} from "../lib/similarity";

type Tab = "results" | "matrix" | "passages";

export function SimilarityScreen() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("Choose at least two documents to begin.");
  const [documents, setDocuments] = useState<SimilarityDocument[]>([]);
  const [results, setResults] = useState<SimilarityResult[]>([]);
  const [tab, setTab] = useState<Tab>("results");
  const [docA, setDocA] = useState(0);
  const [docB, setDocB] = useState(1);
  const [matches, setMatches] = useState<PassageMatch[] | null>(null);
  const [inspecting, setInspecting] = useState(false);

  const totalPairs = files.length > 1 ? files.length * (files.length - 1) / 2 : 0;
  const highest = results[0]?.overall;
  const veryHigh = results.filter((result) => result.overall >= 0.7).length;
  const selectedResult = results.find(
    (result) => (result.i === docA && result.j === docB) || (result.i === docB && result.j === docA)
  );

  const matrix = useMemo(() => {
    const values = Array.from({ length: documents.length }, () => Array<number | null>(documents.length).fill(null));
    for (const result of results) values[result.i]![result.j] = values[result.j]![result.i] = result.overall;
    return values;
  }, [documents.length, results]);

  const acceptFiles = (incoming: File[]) => {
    const supported = incoming.filter((file) => /\.(docx|pdf|txt)$/i.test(file.name));
    setFiles(supported.slice(0, 30));
    setResults([]);
    setDocuments([]);
    setMatches(null);
    setStatus(supported.length < 2 ? "Choose at least two DOCX, PDF, or TXT files." : `${Math.min(supported.length, 30)} documents ready for comparison.`);
  };

  const run = async () => {
    if (files.length < 2) { setStatus("Please select at least two documents."); return; }
    setBusy(true);
    setResults([]);
    setDocuments([]);
    setMatches(null);
    try {
      const output = await compareFiles(files, setStatus);
      setDocuments(output.documents);
      setResults(output.results);
      setDocA(0);
      setDocB(1);
      setTab("results");
      setStatus(`Finished. All ${output.results.length} unique document pairs were checked.`);
    } catch (error) {
      setStatus(`Could not complete comparison: ${error instanceof Error ? error.message : "Unknown error"}`);
    } finally {
      setBusy(false);
    }
  };

  const inspect = async () => {
    if (docA === docB || !documents[docA] || !documents[docB]) return;
    setInspecting(true);
    setMatches(null);
    try { setMatches(await comparePassages(documents[docA]!, documents[docB]!)); }
    finally { setInspecting(false); }
  };

  const openPair = (result: SimilarityResult) => {
    setDocA(result.i);
    setDocB(result.j);
    setMatches(null);
    setTab("passages");
  };

  const downloadCsv = () => {
    const rows = [["Document A", "Document B", "Overall Similarity", "Content Similarity (TF-IDF)", "Phrase Overlap (8-word)", "Similarity Band"]];
    for (const result of results) rows.push([result.a, result.b, percent(result.overall), percent(result.content), percent(result.phrase), band(result.overall).label]);
    const csv = rows.map((row) => row.map((cell) => `"${cell.replace(/"/g, '""')}"`).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "document_similarity_report.csv";
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className="similarity-page">
      <div className="similarity-hero">
        <div>
          <span className="eyebrow">Document intelligence suite</span>
          <h1>Compare documents with clarity—not guesswork.</h1>
          <p>Measure content similarity, direct phrase overlap, and matching passages across every unique pair of DOCX, PDF, and TXT files.</p>
        </div>
        <div className="similarity-formula" aria-label="Similarity score formula">
          <span>Overall similarity</span><strong><b>70%</b> content <i>+</i> <b>30%</b> phrase overlap</strong>
          <small>References and widespread boilerplate are automatically excluded.</small>
        </div>
      </div>

      <div className="similarity-layout">
        <aside className="similarity-sidebar card">
          <div className="tool-card-heading"><span aria-hidden="true">⇄</span><div><strong>Comparison set</strong><small>2–30 documents</small></div></div>
          <div
            className={`similarity-drop${dragging ? " dragging" : ""}`}
            role="button" tabIndex={0}
            onClick={() => inputRef.current?.click()}
            onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") inputRef.current?.click(); }}
            onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => { event.preventDefault(); setDragging(false); acceptFiles([...event.dataTransfer.files]); }}
          >
            <span className="multi-file-icon" aria-hidden="true"><i /><i /><b>+</b></span>
            <strong>Drop documents here</strong><small>or browse from your computer</small>
            <span className="supported">DOCX · PDF · TXT</span>
            <input ref={inputRef} className="sr-only" type="file" multiple accept=".docx,.pdf,.txt" onChange={(event) => acceptFiles([...(event.target.files ?? [])])} />
          </div>

          <div className="selected-files">
            {files.length === 0 && <p>No documents selected yet.</p>}
            {files.map((file, index) => (
              <div key={`${file.name}-${index}`}><span>{file.name.split(".").pop()?.toUpperCase()}</span><p><strong>{file.name}</strong><small>{(file.size / 1024).toFixed(1)} KB</small></p></div>
            ))}
          </div>

          <details className="comparison-settings">
            <summary>How comparison works</summary>
            <ul>
              <li>Reference lists are excluded automatically.</li>
              <li>Capitalization, punctuation, and spacing are normalized.</li>
              <li>TF-IDF uses meaningful words and word pairs.</li>
              <li>Repeated boilerplate is suppressed across large sets.</li>
              <li>Passage review begins at 58% similarity.</li>
            </ul>
          </details>

          <button className="btn primary similarity-run" disabled={busy || files.length < 2} onClick={() => void run()}>
            {busy ? "Comparing documents…" : `Run ${totalPairs || "full"} pair comparison`}
          </button>
          <button className="btn" disabled={!results.length} onClick={downloadCsv}>Download CSV report</button>
          <div className={`comparison-status${busy ? " busy" : ""}`}><span aria-hidden="true" />{status}</div>
          <p className="local-note"><b>Private:</b> comparison runs in your browser. The selected files are not uploaded to the APA formatter server.</p>
        </aside>

        <section className="similarity-workspace">
          <div className="similarity-stats">
            <Metric label="Documents" value={String(files.length)} detail={`${totalPairs} possible pairs`} />
            <Metric label="Pairs checked" value={String(results.length)} detail={results.length ? "100% coverage" : "Waiting to run"} />
            <Metric label="Highest match" value={highest == null ? "—" : percent(highest)} detail={highest == null ? "No result yet" : band(highest).label} tone={highest == null ? undefined : band(highest).className} />
            <Metric label="Very-high pairs" value={String(veryHigh)} detail="70% or above" />
          </div>

          <div className="similarity-results card">
            <div className="workspace-tabs" role="tablist">
              {(["results", "matrix", "passages"] as Tab[]).map((name) => (
                <button key={name} className={tab === name ? "active" : ""} onClick={() => setTab(name)}>
                  {name === "results" ? "Ranked results" : name === "matrix" ? "Similarity matrix" : "Inspect passages"}
                </button>
              ))}
            </div>

            {tab === "results" && <ResultsTable results={results} onOpen={openPair} />}
            {tab === "matrix" && (
              <div className="matrix-scroll">
                {!documents.length ? <EmptyState title="Your matrix will appear here" text="Run a comparison to see every relationship at a glance." /> : (
                  <table className="similarity-matrix"><thead><tr><th>Document</th>{documents.map((_, index) => <th key={index}>{index + 1}</th>)}</tr></thead>
                    <tbody>{documents.map((document, i) => <tr key={document.name}><th>{i + 1}. {document.name}</th>{documents.map((_, j) => {
                      const value = matrix[i]![j];
                      return <td key={j} className={value == null ? "diagonal" : ""} style={value == null ? undefined : { "--match": Math.max(.05, value) } as CSSProperties} onClick={() => { if (value != null) openPair(results.find((result) => (result.i === i && result.j === j) || (result.i === j && result.j === i))!); }}>{value == null ? "—" : percent(value)}</td>;
                    })}</tr>)}</tbody>
                  </table>
                )}
              </div>
            )}
            {tab === "passages" && (
              <div className="passage-inspector">
                {!documents.length ? <EmptyState title="Inspect matching passages" text="Complete a comparison, then choose any two documents." /> : <>
                  <div className="pair-picker">
                    <label>Document A<select value={docA} onChange={(event) => { setDocA(Number(event.target.value)); setMatches(null); }}>{documents.map((document, index) => <option value={index} key={document.name}>{document.name}</option>)}</select></label>
                    <span aria-hidden="true">⇄</span>
                    <label>Document B<select value={docB} onChange={(event) => { setDocB(Number(event.target.value)); setMatches(null); }}>{documents.map((document, index) => <option value={index} key={document.name}>{document.name}</option>)}</select></label>
                    <button className="btn primary" disabled={docA === docB || inspecting} onClick={() => void inspect()}>{inspecting ? "Inspecting…" : "Find matching passages"}</button>
                  </div>
                  {docA === docB && <div className="inline-notice">Choose two different documents.</div>}
                  {selectedResult && <div className="pair-score"><span><small>Overall</small><strong>{percent(selectedResult.overall)}</strong></span><span><small>Content</small><strong>{percent(selectedResult.content)}</strong></span><span><small>Phrase overlap</small><strong>{percent(selectedResult.phrase)}</strong></span><span className={`similarity-band ${band(selectedResult.overall).className}`}>{band(selectedResult.overall).label}</span></div>}
                  {matches?.length === 0 && <div className="inline-notice">No passage pair reached the automatic 58% review threshold.</div>}
                  {matches?.map((match, index) => <div className="passage-match" key={index}><div><b>Match {index + 1}</b><span>{percent(match.score)} passage similarity{match.shared ? ` · ${match.shared} shared phrases` : ""}</span></div><blockquote><strong>{documents[docA]!.name}</strong>{match.a}</blockquote><blockquote><strong>{documents[docB]!.name}</strong>{match.b}</blockquote></div>)}
                </>}
              </div>
            )}
          </div>
        </section>
      </div>
    </div>
  );
}

function Metric(props: { label: string; value: string; detail: string; tone?: string }) {
  return <div className="similarity-metric"><span>{props.label}</span><strong>{props.value}</strong><small className={props.tone}>{props.detail}</small></div>;
}

function ResultsTable(props: { results: SimilarityResult[]; onOpen: (result: SimilarityResult) => void }) {
  if (!props.results.length) return <EmptyState title="Results will be ranked here" text="Each unique document pair is scored exactly once." />;
  return <div className="results-table-wrap"><table className="results-table"><thead><tr><th>Document pair</th><th>Overall</th><th>Content</th><th>Phrase overlap</th><th>Band</th><th /></tr></thead><tbody>
    {props.results.map((result) => { const scoreBand = band(result.overall); return <tr key={`${result.i}-${result.j}`}><td><strong>{result.a}</strong><span>{result.b}</span></td><td><b>{percent(result.overall)}</b></td><td>{percent(result.content)}</td><td>{percent(result.phrase)}</td><td><span className={`similarity-band ${scoreBand.className}`}>{scoreBand.label}</span></td><td><button aria-label={`Inspect ${result.a} and ${result.b}`} onClick={() => props.onOpen(result)}>→</button></td></tr>; })}
  </tbody></table></div>;
}

function EmptyState(props: { title: string; text: string }) {
  return <div className="tool-empty"><span aria-hidden="true">◎</span><strong>{props.title}</strong><p>{props.text}</p></div>;
}
