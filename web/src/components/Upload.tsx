import { Fragment, useRef, useState } from "react";
import { uploadDocument, type UploadResponse } from "../lib/api";

export function UploadScreen(props: {
  onUploaded: (session: UploadResponse) => void;
  onError: (message: string) => void;
  batchMode?: boolean;
  onBatchModeChange?: (batch: boolean) => void;
}) {
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const accept = async (file: File | undefined) => {
    if (!file || busy) return;
    if (!file.name.toLowerCase().endsWith(".docx")) {
      props.onError("This version currently supports Microsoft Word .docx files.");
      return;
    }
    setBusy(true);
    try {
      const session = await uploadDocument(file);
      props.onUploaded(session);
    } catch (e) {
      props.onError(e instanceof Error ? e.message : "Upload failed.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <Fragment>
      {props.onBatchModeChange && (
        <div className="formatter-mode-toggle" role="tablist" aria-label="Choose how many documents to process">
          <button
            role="tab"
            aria-selected={!props.batchMode}
            className={!props.batchMode ? "active" : ""}
            onClick={() => props.onBatchModeChange?.(false)}
          >
            Single document
          </button>
          <button
            role="tab"
            aria-selected={!!props.batchMode}
            className={props.batchMode ? "active" : ""}
            onClick={() => props.onBatchModeChange?.(true)}
          >
            Multiple documents
          </button>
        </div>
      )}
      <section className="landing">
      <div className="hero-copy">
        <span className="eyebrow"><span aria-hidden="true">✦</span> Built for serious academic work</span>
        <h1>Turn a rough Word paper into a <em>submission-ready</em> APA 7 document.</h1>
        <p className="hero-lede">
          A precise formatter and auditor for headings, citations, references,
          title pages, tables, figures, spacing, and document structure.
        </p>
        <div className="hero-proof" aria-label="Product guarantees">
          <span><b>✓</b> Wording preserved</span>
          <span><b>✓</b> Original untouched</span>
          <span><b>✓</b> Detailed audit trail</span>
        </div>
        <div className="mini-preview" aria-hidden="true">
          <div className="preview-paper">
            <span className="preview-page">1</span>
            <i className="preview-title" />
            <i /><i /><i className="short" />
            <strong>Heading</strong>
            <i /><i /><i />
          </div>
          <div className="preview-score">
            <span>APA readiness</span>
            <strong>96<small>%</small></strong>
            <div><i /></div>
            <small>Formatting corrected · 3 items to review</small>
          </div>
        </div>
      </div>

      <div className="upload-panel">
        <div className="upload-panel-head">
          <span className="upload-icon" aria-hidden="true">↥</span>
          <div><strong>Start with your document</strong><span>Upload a Microsoft Word file</span></div>
        </div>
        <div
          className={`dropzone${drag ? " drag" : ""}${busy ? " busy" : ""}`}
          role="button"
          tabIndex={0}
          aria-label="Upload a Word .docx document by clicking or dropping a file"
          onClick={() => inputRef.current?.click()}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              inputRef.current?.click();
            }
          }}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); void accept(e.dataTransfer.files?.[0]); }}
        >
          <div className="file-orbit" aria-hidden="true"><span>W</span></div>
          <p><strong>{busy ? "Uploading your paper…" : "Drop your paper here"}</strong></p>
          <p className="drop-or">or choose it from your computer</p>
          <span className="btn primary upload-cta" aria-hidden="true">
            {busy ? "Please wait…" : "Select Word document"}
          </span>
          <p className="hint">DOCX · up to 25 MB</p>
          <input
            ref={inputRef}
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            className="sr-only"
            aria-label="Choose a Word document"
            onChange={(e) => void accept(e.target.files?.[0] ?? undefined)}
          />
        </div>
        <div className="privacy-strip"><span aria-hidden="true">◇</span><p><strong>Private by design</strong>Your file is never used to train AI.</p></div>
      </div>

      <div className="capability-strip">
        <div><span>01</span><p><strong>Format intelligently</strong>APA typography, headings, tables, figures, and pagination.</p></div>
        <div><span>02</span><p><strong>Verify precisely</strong>Citations and references checked with a transparent report.</p></div>
        <div><span>03</span><p><strong>Download confidently</strong>A corrected DOCX with every applied change documented.</p></div>
      </div>
      </section>
    </Fragment>
  );
}
