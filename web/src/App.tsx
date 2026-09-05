import { lazy, Suspense, useCallback, useEffect, useRef, useState } from "react";
import {
  getReport,
  getStatus,
  startProcessing,
  type ProcessSettings,
  type ReportResponse,
  type SessionInfo,
  type UploadResponse,
} from "./lib/api";
import { UploadScreen } from "./components/Upload";
import { ConfigureScreen } from "./components/Configure";
import { ProcessingScreen } from "./components/Processing";
import { ResultsScreen } from "./components/Results";
const SimilarityScreen = lazy(() =>
  import("./components/Similarity").then((module) => ({ default: module.SimilarityScreen }))
);
const MergeScreen = lazy(() =>
  import("./components/Merge").then((module) => ({ default: module.MergeScreen }))
);
const TurnitinScreen = lazy(() =>
  import("./components/Turnitin").then((module) => ({ default: module.TurnitinScreen }))
);
const BatchWorkspace = lazy(() =>
  import("./components/BatchWorkspace").then((module) => ({ default: module.BatchWorkspace }))
);

const PRODUCT_NAME = "APA 7 Document Auditor"; // configurable product name

type Screen =
  | { kind: "upload" }
  | { kind: "configure"; session: UploadResponse }
  | { kind: "processing"; session: UploadResponse; status: SessionInfo | null }
  | { kind: "results"; session: UploadResponse; report: ReportResponse };

type Tool = "formatter" | "similarity" | "merger" | "turnitin";
type Theme = "light" | "dark";

function initialTheme(): Theme {
  const saved = window.localStorage.getItem("scholarlyworks-theme");
  if (saved === "light" || saved === "dark") return saved;
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function toolFromPath(): Tool {
  const path = window.location.pathname.replace(/\/+$/, "") || "/";
  if (path === "/turnitin") return "turnitin";
  if (path === "/similarity") return "similarity";
  if (path === "/merger") return "merger";
  return "formatter";
}

const TOOL_PATH: Record<Tool, string> = {
  formatter: "/",
  similarity: "/similarity",
  merger: "/merger",
  turnitin: "/turnitin",
};

export function App() {
  const [tool, setTool] = useState<Tool>(toolFromPath);
  const [theme, setTheme] = useState<Theme>(initialTheme);
  const [formatterMode, setFormatterMode] = useState<"single" | "batch">("single");
  const [screen, setScreen] = useState<Screen>({ kind: "upload" });
  const [error, setError] = useState<string | null>(null);
  const pollRef = useRef<number | null>(null);

  const stopPolling = useCallback(() => {
    if (pollRef.current != null) {
      window.clearInterval(pollRef.current);
      pollRef.current = null;
    }
  }, []);
  useEffect(() => stopPolling, [stopPolling]);
  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.documentElement.style.colorScheme = theme;
    window.localStorage.setItem("scholarlyworks-theme", theme);
  }, [theme]);
  useEffect(() => {
    const onPopState = () => { setTool(toolFromPath()); setError(null); };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  const beginPolling = useCallback(
    (session: UploadResponse) => {
      stopPolling();
      pollRef.current = window.setInterval(async () => {
        try {
          const status = await getStatus(session.id);
          if (status.status === "ready") {
            stopPolling();
            const report = await getReport(session.id);
            setScreen({ kind: "results", session, report });
          } else if (status.status === "error") {
            stopPolling();
            setError(status.error ?? "Processing failed.");
            setScreen({ kind: "configure", session });
          } else {
            setScreen({ kind: "processing", session, status });
          }
        } catch (e) {
          stopPolling();
          setError(e instanceof Error ? e.message : "Connection lost.");
          setScreen({ kind: "configure", session });
        }
      }, 700);
    },
    [stopPolling]
  );

  const handleUploaded = (session: UploadResponse) => {
    setError(null);
    setScreen({ kind: "configure", session });
  };

  const handleStart = async (session: UploadResponse, settings: ProcessSettings) => {
    setError(null);
    try {
      await startProcessing(session.id, settings);
      setScreen({ kind: "processing", session, status: null });
      beginPolling(session);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not start processing.");
    }
  };

  const handleReportUpdate = (report: ReportResponse) => {
    if (screen.kind === "results") {
      setScreen({ ...screen, report });
    }
  };

  const handleRegenerate = (session: UploadResponse) => {
    setScreen({ kind: "processing", session, status: null });
    beginPolling(session);
  };

  const reset = () => {
    stopPolling();
    setError(null);
    setScreen({ kind: "upload" });
  };

  const selectTool = (next: Tool) => {
    setTool(next);
    if (window.location.pathname !== TOOL_PATH[next]) window.history.pushState({}, "", TOOL_PATH[next]);
    if (next === "formatter") reset();
    setError(null);
  };

  const step = screen.kind === "upload" ? 1 : screen.kind === "configure" ? 2 : 3;

  return (
    <div className="app-frame">
      <div className="ambient ambient-one" aria-hidden="true" />
      <div className="ambient ambient-two" aria-hidden="true" />
      <div className="shell">
        <header className="topbar">
          <button className="brand-lockup" onClick={() => selectTool("formatter")} aria-label="Return to APA formatter">
            <span className="brand-mark" aria-hidden="true">A7</span>
            <span className="brand-copy">
              <strong>{PRODUCT_NAME}</strong>
              <small>Academic formatting, made dependable</small>
            </span>
          </button>
          <nav className="product-switch" aria-label="Choose a document tool">
            <button className={tool === "formatter" ? "active" : ""} onClick={() => selectTool("formatter")}>APA formatter</button>
            <button className={tool === "similarity" ? "active" : ""} onClick={() => selectTool("similarity")}>Similarity checker</button>
            <button className={tool === "merger" ? "active" : ""} onClick={() => selectTool("merger")}>Document merger</button>
            <button className={tool === "turnitin" ? "active" : ""} onClick={() => selectTool("turnitin")}>Turnitin checker</button>
          </nav>
          <div className="topbar-actions">
            <button
              className="theme-toggle"
              type="button"
              aria-label={`Switch to ${theme === "light" ? "dark" : "light"} mode`}
              aria-pressed={theme === "dark"}
              onClick={() => setTheme((current) => current === "light" ? "dark" : "light")}
            >
              <span aria-hidden="true">{theme === "light" ? "◐" : "☀"}</span>
              {theme === "light" ? "Dark" : "Light"}
            </button>
            <span className="secure-pill"><span aria-hidden="true">●</span> Private processing</span>
            <span className="edition-pill">APA 7th Edition</span>
          </div>
        </header>

        {tool === "formatter" && formatterMode === "single" ? <nav className="journey" aria-label="Document workflow">
          {["Upload", "Customize", "Review & download"].map((label, index) => {
            const n = index + 1;
            return (
              <div className={`journey-step${step === n ? " active" : ""}${step > n ? " complete" : ""}`} key={label}>
                <span>{step > n ? "✓" : n}</span>
                <strong>{label}</strong>
              </div>
            );
          })}
        </nav> : tool === "formatter" ? <div className="similarity-journey batch-journey"><span>▤</span><strong>Batch APA formatting</strong><small>Same settings for every file · independent progress and downloads</small></div> : tool === "similarity" ? <div className="similarity-journey"><span>⇄</span><strong>Document-to-document comparison</strong><small>Every unique pair · local browser analysis · exportable report</small></div> : tool === "merger" ? <div className="similarity-journey merge-journey"><span>↧</span><strong>Multi-document submission builder</strong><small>Original formatting preserved · reference lists removed · adjustable appendix budget</small></div> : <div className="similarity-journey checker-journey"><span>✓</span><strong>Turnitin checker</strong><small>Official similarity and AI results · secure backend processing</small></div>}

        {error && (
          <div className="error-box" role="alert">
            <strong>We couldn’t continue.</strong>
            <span>{error}</span>
          </div>
        )}

        <main>
          {tool === "similarity" ? <Suspense fallback={<div className="tool-loader"><span>A7</span><p>Preparing the comparison workspace…</p></div>}><SimilarityScreen /></Suspense> : tool === "merger" ? <Suspense fallback={<div className="tool-loader"><span>A7</span><p>Preparing the merge workspace…</p></div>}><MergeScreen /></Suspense> : tool === "turnitin" ? <Suspense fallback={<div className="tool-loader"><span>A7</span><p>Preparing the checker…</p></div>}><TurnitinScreen /></Suspense> : tool === "formatter" && formatterMode === "batch" ? <Suspense fallback={<div className="tool-loader"><span>A7</span><p>Preparing the batch workspace…</p></div>}><BatchWorkspace onSwitchToSingle={() => setFormatterMode("single")} /></Suspense> : <>
          {screen.kind === "upload" && (
            <UploadScreen
              onUploaded={handleUploaded}
              onError={setError}
              batchMode={false}
              onBatchModeChange={(batch) => { if (batch) setFormatterMode("batch"); }}
            />
          )}
          {screen.kind === "configure" && (
            <ConfigureScreen
              session={screen.session}
              onStart={(settings) => handleStart(screen.session, settings)}
              onReset={reset}
            />
          )}
          {screen.kind === "processing" && (
            <ProcessingScreen session={screen.session} status={screen.status} />
          )}
          {screen.kind === "results" && (
            <ResultsScreen
              session={screen.session}
              data={screen.report}
              onReportUpdate={handleReportUpdate}
              onRegenerate={() => handleRegenerate(screen.session)}
              onReset={reset}
            />
          )}
          </>}
        </main>

        <footer className="site-footer">
          <div>
            <strong>Your work stays yours.</strong>
            <span> Files are processed only for formatting and validation and are never used for training.</span>
          </div>
          <span>Original wording preserved · Original file untouched</span>
        </footer>
      </div>
    </div>
  );
}
