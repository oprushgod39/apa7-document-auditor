import { useEffect, useRef, useState } from "react";
import {
  ApiError,
  downloadUrl,
  getStatus,
  reportDownloadUrl,
  startProcessing,
  uploadDocument,
  type ProcessSettings,
} from "../lib/api";
import { ALL_STAGES } from "./Processing";
import { BatchUpload, type BatchFileEntry } from "./BatchUpload";
import { BatchConfigureScreen } from "./BatchConfigure";

// Bounded concurrency: enough to keep several documents moving at once
// without hammering the backend's shared session store or the Crossref
// verification rate limit (CROSSREF_MAX_REQUESTS_PER_RUN) — and, more
// importantly, without blowing through the API's global rate limiter
// (RATE_LIMIT_MAX requests per RATE_LIMIT_WINDOW_MS, applied per IP across
// every /api route). Each in-flight file polls its own status roughly every
// 700ms, so even 3 concurrent files can occasionally trip the limiter during
// slower processing — callApi() below retries RATE_LIMITED responses with
// backoff rather than failing the row, so a burst just slows down instead
// of erroring out.
const BATCH_CONCURRENCY = 3;
const POLL_INTERVAL_MS = 700;
const STAGE_LABELS = new Map(ALL_STAGES);

type RowStatus = "queued" | "uploading" | "processing" | "ready" | "error";

interface BatchRow {
  id: string;
  file: File;
  status: RowStatus;
  sessionId?: string;
  stageLabel?: string;
  error?: string;
  /** Set once uploaded: true when this file has neither an existing title
   * page nor a detectable title, so no title page will be auto-created for
   * it (the app never invents a title — see BatchConfigureScreen's note). */
  noTitleDetected?: boolean;
}

type Step = "select" | "configure" | "run";

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Retries a call a bounded number of times when the server responds with
 * RATE_LIMITED, using backoff. Any other error (including a genuine
 * processing failure) is thrown immediately so the row can show it. */
async function callApi<T>(fn: () => Promise<T>, cancelled: () => boolean): Promise<T> {
  let attempt = 0;
  for (;;) {
    if (cancelled()) throw new Error("cancelled");
    try {
      return await fn();
    } catch (e) {
      attempt += 1;
      if (e instanceof ApiError && e.code === "RATE_LIMITED" && attempt < 20) {
        await sleep(Math.min(1000 * attempt, 6000));
        continue;
      }
      throw e;
    }
  }
}

function currentStageLabel(stages: { key: string; status: string }[]): string {
  const running = stages.find((s) => s.status === "running");
  if (running) return STAGE_LABELS.get(running.key) ?? "Processing…";
  const lastDone = [...stages].reverse().find((s) => s.status === "done");
  if (lastDone) return `${STAGE_LABELS.get(lastDone.key) ?? "Processing"} — complete`;
  return "Waiting to start…";
}

export function BatchWorkspace(props: { onSwitchToSingle: () => void }) {
  const [step, setStep] = useState<Step>("select");
  const [pendingFiles, setPendingFiles] = useState<BatchFileEntry[]>([]);
  const [rows, setRows] = useState<BatchRow[]>([]);
  const [settings, setSettings] = useState<ProcessSettings | null>(null);
  const cancelledRef = useRef(false);

  useEffect(() => {
    cancelledRef.current = false;
    return () => {
      cancelledRef.current = true;
    };
  }, []);

  const updateRow = (id: string, patch: Partial<BatchRow>) => {
    if (cancelledRef.current) return;
    setRows((current) => current.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  };

  const runFile = async (row: BatchRow, activeSettings: ProcessSettings) => {
    const cancelled = () => cancelledRef.current;
    try {
      updateRow(row.id, { status: "uploading" });
      const session = await callApi(() => uploadDocument(row.file), cancelled);
      if (cancelled()) return;

      // Title and author always come from this specific document — never from
      // the batch-shared settings — matching the single-file flow and the
      // "never invent" title-page rule. If this file has no title page and no
      // detectable title, no title page will be auto-created for it; flag
      // that up front so it's visible even before processing finishes.
      //
      // Author falls back to the same "Name" placeholder the single-file
      // Configure screen defaults to when nothing is detected — omitting it
      // entirely (as an earlier version of this code did) silently drops the
      // author line from the generated title page instead of showing a
      // placeholder for the student to fill in.
      const detected = session.detected.metadata;
      const noTitleDetected = !session.detected.hasTitlePage && !detected.title;
      const perFileMetadata: Record<string, string> = { ...activeSettings.metadata };
      if (detected.title) perFileMetadata.title = detected.title;
      perFileMetadata.author = detected.author || "Name";
      const perFileSettings: ProcessSettings = { ...activeSettings, metadata: perFileMetadata };

      updateRow(row.id, {
        sessionId: session.id,
        status: "processing",
        stageLabel: "Starting…",
        noTitleDetected,
      });

      await callApi(() => startProcessing(session.id, perFileSettings), cancelled);
      if (cancelled()) return;

      for (;;) {
        if (cancelled()) return;
        await sleep(POLL_INTERVAL_MS);
        if (cancelled()) return;
        const status = await callApi(() => getStatus(session.id), cancelled);
        if (cancelled()) return;
        if (status.status === "ready") {
          updateRow(row.id, { status: "ready", stageLabel: undefined });
          return;
        }
        if (status.status === "error") {
          updateRow(row.id, { status: "error", error: status.error ?? "Processing failed." });
          return;
        }
        updateRow(row.id, { stageLabel: currentStageLabel(status.stages) });
      }
    } catch (e) {
      if (cancelled()) return;
      const message = e instanceof ApiError ? e.message : e instanceof Error ? e.message : "Something went wrong.";
      updateRow(row.id, { status: "error", error: message });
    }
  };

  const startBatch = (files: BatchFileEntry[], activeSettings: ProcessSettings) => {
    const initialRows: BatchRow[] = files.map((f) => ({ id: f.id, file: f.file, status: "queued" }));
    setRows(initialRows);
    setSettings(activeSettings);
    setStep("run");

    // Bounded-concurrency worker pool: a fixed number of "lanes" pull the
    // next queued file as soon as they free up, rather than firing every
    // upload/process call at once.
    let cursor = 0;
    const lanes = Array.from({ length: Math.min(BATCH_CONCURRENCY, initialRows.length) }, async () => {
      for (;;) {
        if (cancelledRef.current) return;
        const index = cursor;
        cursor += 1;
        if (index >= initialRows.length) return;
        await runFile(initialRows[index]!, activeSettings);
      }
    });
    void Promise.all(lanes);
  };

  const reset = () => {
    setRows([]);
    setSettings(null);
    setPendingFiles([]);
    setStep("select");
  };

  if (step === "select") {
    return (
      <BatchUpload
        onContinue={(files) => {
          setStep("configure");
          // stash selection until settings are confirmed
          setPendingFiles(files);
        }}
        onSwitchToSingle={props.onSwitchToSingle}
      />
    );
  }

  if (step === "configure") {
    return (
      <BatchConfigureScreen
        files={pendingFiles}
        onStart={(s) => startBatch(pendingFiles, s)}
        onBack={() => setStep("select")}
      />
    );
  }

  const ready = rows.filter((r) => r.status === "ready").length;
  const errored = rows.filter((r) => r.status === "error").length;
  const inProgress = rows.length - ready - errored;
  const canDownloadDocx = settings?.mode !== "check";

  return (
    <section className="batch-run-page">
      <div className="page-heading">
        <div>
          <span className="eyebrow">Batch results</span>
          <h1>Processing {rows.length} document{rows.length === 1 ? "" : "s"}</h1>
        </div>
        <button className="btn" onClick={reset}>
          Start a new batch
        </button>
      </div>

      <div className="card batch-summary-card">
        <div className="batch-summary">
          <span>
            <strong>{ready}</strong> of {rows.length} ready
          </span>
          {errored > 0 && (
            <span className="err">
              <strong>{errored}</strong> failed
            </span>
          )}
          {inProgress > 0 && (
            <span>
              <strong>{inProgress}</strong> processing
            </span>
          )}
        </div>
      </div>

      <div className="card batch-table-card">
        <div className="results-table-wrap">
          <table className="results-table batch-table">
            <thead>
              <tr>
                <th>Document</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <BatchTableRow key={row.id} row={row} canDownloadDocx={canDownloadDocx} />
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}

function BatchTableRow(props: { row: BatchRow; canDownloadDocx: boolean }) {
  const { row } = props;
  return (
    <tr>
      <td>
        <strong>{row.file.name}</strong>
        <span>{(row.file.size / 1024 / 1024).toFixed(2)} MB</span>
      </td>
      <td>
        {row.status === "ready" && <span className="badge pass">READY</span>}
        {row.status === "error" && <span className="badge fail">ERROR</span>}
        {(row.status === "queued" || row.status === "uploading" || row.status === "processing") && (
          <span className="batch-status-live">
            <span className="batch-status-dot" aria-hidden="true" />
            {row.status === "queued" ? "Queued…" : row.status === "uploading" ? "Uploading…" : row.stageLabel ?? "Processing…"}
          </span>
        )}
        {row.status === "error" && row.error && <div className="batch-row-error">{row.error}</div>}
        {row.noTitleDetected && row.status !== "error" && (
          <div className="batch-row-warning">
            No title detected — title page not created. Format this file
            individually to type in a title.
          </div>
        )}
      </td>
      <td>
        {row.status === "ready" && row.sessionId && (
          <div className="batch-row-actions">
            {props.canDownloadDocx && (
              <a className="btn small primary" href={downloadUrl(row.sessionId)}>
                Download
              </a>
            )}
            <a className="btn small" href={reportDownloadUrl(row.sessionId)}>
              Report
            </a>
          </div>
        )}
      </td>
    </tr>
  );
}
