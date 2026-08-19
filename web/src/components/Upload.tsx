import { useRef, useState } from "react";
import { uploadDocument, type UploadResponse } from "../lib/api";

export function UploadScreen(props: {
  onUploaded: (session: UploadResponse) => void;
  onError: (message: string) => void;
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
    <section className="hero">
      <h1>Upload your Word paper and format it according to APA 7.</h1>
      <p>
        Checks formatting, headings, citations, references, tables, figures,
        and other APA requirements — then produces a corrected .docx and a
        detailed compliance report. Your wording is never rewritten.
      </p>
      <div
        className={`dropzone${drag ? " drag" : ""}`}
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
        onDragOver={(e) => {
          e.preventDefault();
          setDrag(true);
        }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDrag(false);
          void accept(e.dataTransfer.files?.[0]);
        }}
      >
        <p>
          <strong>{busy ? "Uploading…" : "Drop your .docx here"}</strong>
        </p>
        <p>or</p>
        <span className="btn primary" aria-hidden="true">
          Choose Document
        </span>
        <p className="hint">
          .docx only · maximum 25 MB · the original file is never modified
        </p>
        <input
          ref={inputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="sr-only"
          aria-label="Choose a Word document"
          onChange={(e) => void accept(e.target.files?.[0] ?? undefined)}
        />
      </div>
    </section>
  );
}
