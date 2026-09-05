import { useEffect, useRef, useState } from "react";
import {
  checkerAvailability,
  checkStatus,
  createCheck,
  downloadCheckerOriginal,
  downloadCheckerReport,
  unlockAiReport,
  viewCheckerReport,
  type CheckerAvailability,
  type CheckerResult,
} from "../lib/similarityscan";

type Busy = "upload" | "restore" | "view-similarity" | "view-ai" | "download-similarity" | "download-ai" | "download-original" | "unlock" | null;
const ACCEPTED = /\.(pdf|docx|txt|rtf|xlsx)$/i;
const HISTORY_KEY = "scholarlyworks-similarityscan-history-v1";

interface SavedCheck {
  id: string;
  accessToken: string;
  orderId: string;
  filename: string;
  status: string;
  savedAt: number;
}

function readSavedChecks(): SavedCheck[] {
  try {
    const raw = window.localStorage.getItem(HISTORY_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item): item is SavedCheck => Boolean(
      item && typeof item === "object" && typeof (item as SavedCheck).id === "string" &&
      typeof (item as SavedCheck).accessToken === "string" && typeof (item as SavedCheck).filename === "string"
    ));
  } catch { return []; }
}

function writeSavedChecks(checks: SavedCheck[]): void {
  try { window.localStorage.setItem(HISTORY_KEY, JSON.stringify(checks)); } catch { /* Browser storage may be disabled. */ }
}

function savedFrom(result: CheckerResult, accessToken: string): SavedCheck {
  return { id: result.id, accessToken, orderId: result.orderId, filename: result.filename, status: result.status, savedAt: Date.now() };
}

function displayDate(timestamp: number): string {
  try { return new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(timestamp); }
  catch { return "Saved check"; }
}

export function TurnitinScreen() {
  const [availability, setAvailability] = useState<CheckerAvailability | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<CheckerResult | null>(null);
  const [accessToken, setAccessToken] = useState("");
  const [savedChecks, setSavedChecks] = useState<SavedCheck[]>([]);
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const refreshAvailability = () => {
    setError("");
    checkerAvailability().then(setAvailability).catch((cause) => setError(cause instanceof Error ? cause.message : "The checker is unavailable."));
  };

  const remember = (nextResult: CheckerResult, nextToken: string) => {
    if (!nextToken) return;
    const next = [savedFrom(nextResult, nextToken), ...readSavedChecks().filter((entry) => entry.id !== nextResult.id)];
    writeSavedChecks(next);
    setSavedChecks(next);
  };

  const restore = async (saved: SavedCheck) => {
    setBusy("restore"); setError("");
    try {
      const next = await checkStatus(saved.id, saved.accessToken);
      setAccessToken(saved.accessToken);
      setResult(next);
      remember(next, saved.accessToken);
    } catch (cause) {
      setError(cause instanceof Error ? `${cause.message} This saved check may no longer be available.` : "This saved check could not be restored.");
    } finally { setBusy(null); }
  };

  useEffect(() => {
    refreshAvailability();
    const saved = readSavedChecks();
    setSavedChecks(saved);
    if (saved[0]) void restore(saved[0]);
  }, []);

  useEffect(() => {
    if (!result || !accessToken || result.terminal) return;
    let stopped = false;
    const timer = window.setInterval(async () => {
      try {
        const next = await checkStatus(result.id, accessToken);
        if (!stopped) { setResult(next); setError(""); }
      } catch (cause) {
        if (!stopped) setError(cause instanceof Error ? cause.message : "Status could not be refreshed. Retrying…");
      }
    }, 5000);
    return () => { stopped = true; window.clearInterval(timer); };
  }, [result?.id, result?.terminal, accessToken]);

  useEffect(() => {
    if (result && accessToken) remember(result, accessToken);
  }, [result, accessToken]);

  const choose = (candidate: File | null) => {
    if (!candidate) return;
    if (!ACCEPTED.test(candidate.name)) { setFile(null); setError("Choose a PDF, DOCX, TXT, RTF, or XLSX file."); return; }
    if (availability && candidate.size > availability.maxUploadBytes) {
      setFile(null); setError(`The maximum file size is ${Math.round(availability.maxUploadBytes / 1024 / 1024)} MB.`); return;
    }
    setFile(candidate); setError("");
  };

  const start = async () => {
    if (!file || !availability?.configured || !availability.ready || !availability.acceptingUploads) return;
    setBusy("upload"); setError("");
    try {
      const created = await createCheck(file, availability.persistenceEnabled);
      setAccessToken(created.accessToken ?? "");
      setResult(created);
    } catch (cause) { setError(cause instanceof Error ? cause.message : "The document could not be checked."); }
    finally { setBusy(null); }
  };

  const reportAction = async (kind: "similarity" | "ai", mode: "view" | "download") => {
    if (!result || !accessToken) return;
    const state = `${mode}-${kind}` as Busy;
    setBusy(state); setError("");
    try {
      if (mode === "view") await viewCheckerReport(result.id, accessToken, kind);
      else await downloadCheckerReport(result.id, accessToken, kind);
    } catch (cause) { setError(cause instanceof Error ? cause.message : "The report is unavailable."); }
    finally { setBusy(null); }
  };

  const downloadOriginal = async () => {
    if (!result || !accessToken) return;
    setBusy("download-original"); setError("");
    try { await downloadCheckerOriginal(result.id, accessToken, result.filename); }
    catch (cause) { setError(cause instanceof Error ? cause.message : "The original file is unavailable."); }
    finally { setBusy(null); }
  };

  const unlock = async () => {
    if (!result || !accessToken) return;
    setBusy("unlock"); setError("");
    try { setResult(await unlockAiReport(result.id, accessToken)); }
    catch (cause) { setError(cause instanceof Error ? cause.message : "The AI report could not be unlocked."); }
    finally { setBusy(null); }
  };

  const forget = (id: string) => {
    const next = readSavedChecks().filter((entry) => entry.id !== id);
    writeSavedChecks(next); setSavedChecks(next);
  };
  const clearSaved = () => { writeSavedChecks([]); setSavedChecks([]); };
  const reset = () => { setFile(null); setResult(null); setAccessToken(""); setBusy(null); setError(""); refreshAvailability(); };
  const completed = result?.status === "completed";
  const failed = result?.status === "error" || result?.status === "failed_invalid" || result?.status === "cancelled";

  return (
    <section className="checker-page">
      <header className="checker-heading">
        <div><span className="eyebrow">Turnitin checker</span><h1>Upload. Check.<br/><em>See the real result.</em></h1></div>
        <div className={`checker-availability ${availability?.ready ? "ready" : "offline"}`}>
          <span aria-hidden="true" />
          <div><strong>{availability?.ready ? "Checker available" : "Checker unavailable"}</strong><small>{availability?.message || "Checking availability…"}</small></div>
          {availability && !availability.ready && <button onClick={refreshAvailability}>Retry</button>}
        </div>
      </header>

      {error && <div className="checker-error" role="alert">{error}</div>}

      {!result ? <>
        <div className="checker-card checker-upload-card">
          <button className={`checker-drop${file ? " selected" : ""}`} onClick={() => inputRef.current?.click()} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); choose(event.dataTransfer.files[0] ?? null); }}>
            <span>↑</span><strong>{file ? file.name : "Drop your document here"}</strong><small>{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB · Click to replace` : "PDF, DOCX, TXT, RTF, or XLSX"}</small>
          </button>
          <input ref={inputRef} hidden type="file" accept=".pdf,.docx,.txt,.rtf,.xlsx" onChange={(event) => choose(event.target.files?.[0] ?? null)} />
          <div className="checker-upload-footer">
            <p><strong>Secure backend check</strong><span>{availability?.persistenceEnabled ? "Checks are saved here, so you can reopen them after refreshing." : "The provider token is never placed in the website code or browser."}</span></p>
            <button className="btn primary btn-large" disabled={!file || !availability?.ready || !availability.acceptingUploads || busy === "upload" || busy === "restore"} onClick={start}>{busy === "upload" ? "Uploading securely…" : "Check document"}</button>
          </div>
        </div>
        {savedChecks.length > 0 && <aside className="checker-history" aria-label="Saved checker results">
          <div className="checker-history-heading"><div><span className="eyebrow">Saved checks</span><h2>Reopen a previous result</h2></div><button className="checker-text-button" onClick={clearSaved}>Clear this device</button></div>
          <p>Saved only in this browser. Choose a document to refresh its latest official status and report links.</p>
          <div className="checker-history-list">{savedChecks.map((saved) => <div className="checker-history-item" key={saved.id}>
            <button className="checker-history-open" disabled={busy !== null} onClick={() => void restore(saved)}><strong>{saved.filename}</strong><span>{saved.status.replace(/_/g, " ")} · {displayDate(saved.savedAt)}</span></button>
            <button className="checker-history-remove" aria-label={`Forget ${saved.filename}`} onClick={() => forget(saved.id)}>×</button>
          </div>)}</div>
        </aside>}
      </> : <div className="checker-card checker-result-card">
        <div className="checker-file"><div><small>DOCUMENT</small><strong>{result.filename}</strong><span>Order {result.orderId}</span></div><button className="btn small" onClick={reset}>Check another</button></div>

        {!completed && !failed ? <div className="checker-processing" aria-live="polite">
          <div className="checker-loader"><span /></div>
          <h2>{result.status === "queued" ? "Waiting for a checker" : "Processing your document"}</h2>
          <p>The official status is refreshed every 5 seconds. You can safely refresh this page—the check will resume.</p>
          <div className="checker-progress"><span style={{ width: `${result.progress}%` }} /></div>
        </div> : failed ? <div className="checker-processing checker-failed">
          <h2>{result.status === "failed_invalid" ? "This file could not be checked" : "The check did not complete"}</h2>
          <p>{result.message || (result.status === "failed_invalid" ? "The file may be too short, unreadable, damaged, or unsupported." : "Please try again later.")}</p>
          <button className="btn primary" onClick={reset}>Try another file</button>
        </div> : <>
          <div className="checker-scores">
            <button disabled={!result.similarityReportAvailable || busy !== null} onClick={() => reportAction("similarity", "view")}>
              <small>SIMILARITY</small><strong>{result.similarityPercentage == null ? "—" : `${result.similarityPercentage}%`}</strong><span>{result.similarityReportAvailable ? "View report ↗" : "Report unavailable"}</span>
            </button>
            <button disabled={!result.aiReportAvailable || busy !== null} onClick={() => reportAction("ai", "view")}>
              <small>AI WRITING</small><strong>{result.aiPercentage == null ? "—" : `${result.aiPercentage}%`}</strong><span>{result.aiReportAvailable ? "View report ↗" : result.aiLocked ? "AI report locked" : "AI report not available"}</span>
            </button>
          </div>
          {result.aiLocked && <div className="checker-unlock"><span>AI report is locked for this order.{result.aiUnlockPrice != null ? ` Unlocking costs ${result.aiUnlockPrice} credit${result.aiUnlockPrice === 1 ? "" : "s"}.` : ""}</span><button className="btn" disabled={busy !== null} onClick={unlock}>{busy === "unlock" ? "Unlocking…" : "Unlock AI report"}</button></div>}
          <div className="checker-downloads">
            <button className="btn" disabled={!result.similarityReportAvailable || busy !== null} onClick={() => reportAction("similarity", "download")}>{busy === "download-similarity" ? "Downloading…" : "Download Similarity Report"}</button>
            <button className="btn" disabled={!result.aiReportAvailable || busy !== null} onClick={() => reportAction("ai", "download")}>{busy === "download-ai" ? "Downloading…" : result.aiReportAvailable ? "Download AI Report" : "AI report not available"}</button>
            <button className="btn" disabled={busy !== null} onClick={downloadOriginal}>{busy === "download-original" ? "Downloading…" : "Download Original File"}</button>
          </div>
          <p className="checker-note">This check is saved on this device. Report links are refreshed from the provider before every view or download.</p>
        </>}
      </div>}
    </section>
  );
}
