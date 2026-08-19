import { useCallback, useEffect, useRef, useState } from "react";
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

const PRODUCT_NAME = "APA 7 Document Auditor"; // configurable product name

type Screen =
  | { kind: "upload" }
  | { kind: "configure"; session: UploadResponse }
  | { kind: "processing"; session: UploadResponse; status: SessionInfo | null }
  | { kind: "results"; session: UploadResponse; report: ReportResponse };

export function App() {
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

  return (
    <div className="shell">
      <header className="topbar">
        <div className="brand">
          {PRODUCT_NAME}
          <small>Format and verify Word papers against APA 7th Edition</small>
        </div>
        <p className="privacy-note">
          Your document is processed only to perform formatting and validation
          and is automatically deleted according to the configured retention
          policy. Documents are never used for training.
        </p>
      </header>

      {error && (
        <div className="error-box" role="alert">
          {error}
        </div>
      )}

      <main>
        {screen.kind === "upload" && (
          <UploadScreen onUploaded={handleUploaded} onError={setError} />
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
      </main>
    </div>
  );
}
