import { Fragment, useRef, useState } from "react";

export const MAX_BATCH_FILES = 25;
export const MIN_BATCH_FILES = 2;

export interface BatchFileEntry {
  id: string;
  file: File;
}

export function BatchUpload(props: {
  onContinue: (files: BatchFileEntry[]) => void;
  onSwitchToSingle: () => void;
}) {
  const [items, setItems] = useState<BatchFileEntry[]>([]);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const addFiles = (selected: File[]) => {
    setError(null);
    const docx = selected.filter((file) => /\.docx$/i.test(file.name));
    if (docx.length !== selected.length) {
      setError("Only Microsoft Word .docx files can be added to a batch.");
    }
    setItems((current) => {
      const known = new Set(current.map((item) => `${item.file.name}|${item.file.size}|${item.file.lastModified}`));
      const next = [...current];
      for (const file of docx) {
        const key = `${file.name}|${file.size}|${file.lastModified}`;
        if (!known.has(key) && next.length < MAX_BATCH_FILES) {
          known.add(key);
          next.push({ id: crypto.randomUUID(), file });
        }
      }
      if (current.length + docx.length > MAX_BATCH_FILES) {
        setError(`A batch can contain no more than ${MAX_BATCH_FILES} documents.`);
      }
      return next;
    });
  };

  const remove = (id: string) => setItems((current) => current.filter((item) => item.id !== id));

  const totalSize = items.reduce((sum, item) => sum + item.file.size, 0);
  const canContinue = items.length >= MIN_BATCH_FILES && items.length <= MAX_BATCH_FILES;

  return (
    <Fragment>
      <div className="formatter-mode-toggle" role="tablist" aria-label="Choose how many documents to process">
        <button role="tab" aria-selected={false} onClick={props.onSwitchToSingle}>
          Single document
        </button>
        <button role="tab" aria-selected={true} className="active">
          Multiple documents
        </button>
      </div>

      <section className="batch-upload-page">
        <div className="page-heading">
          <div>
            <span className="eyebrow">Batch submission</span>
            <h1>Format many papers at once.</h1>
          </div>
          <p>
            Upload {MIN_BATCH_FILES}–{MAX_BATCH_FILES} Word documents. The same settings will be
            applied to every file, and each one gets its own download once it's ready.
          </p>
        </div>

        <div className="card batch-upload-card">
          <button
            className={`merge-drop${dragging ? " dragging" : ""}`}
            onClick={() => inputRef.current?.click()}
            onDragOver={(event) => {
              event.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDragging(false);
              addFiles([...event.dataTransfer.files]);
            }}
          >
            <span className="merge-drop-icon">＋</span>
            <strong>Drop DOCX files here</strong>
            <small>
              or click to choose files · {MIN_BATCH_FILES}–{MAX_BATCH_FILES} documents
            </small>
          </button>
          <input
            ref={inputRef}
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            multiple
            hidden
            onChange={(event) => addFiles([...(event.target.files ?? [])])}
          />

          {items.length > 0 && (
            <div className="merge-queue batch-queue">
              <div className="merge-queue-head">
                <div>
                  <span>{items.length}</span>
                  <strong>Documents selected</strong>
                </div>
                <button onClick={() => setItems([])}>Clear all</button>
              </div>
              {items.map((item, index) => (
                <article className="merge-item batch-item" key={item.id}>
                  <span className="merge-order">{String(index + 1).padStart(2, "0")}</span>
                  <div className="merge-file-copy">
                    <strong>{item.file.name}</strong>
                    <small>{(item.file.size / 1024 / 1024).toFixed(2)} MB</small>
                  </div>
                  <div className="merge-controls">
                    <button className="remove" aria-label={`Remove ${item.file.name}`} onClick={() => remove(item.id)}>
                      ×
                    </button>
                  </div>
                </article>
              ))}
            </div>
          )}

          {error && <p className="merge-error" role="alert">{error}</p>}

          <div className="action-dock">
            <div>
              <strong>{items.length} file{items.length === 1 ? "" : "s"} selected</strong>
              <span>{(totalSize / 1024 / 1024).toFixed(2)} MB total · up to 25 MB per file</span>
            </div>
            <button
              className="btn primary btn-large"
              disabled={!canContinue}
              onClick={() => props.onContinue(items)}
            >
              Continue to settings <span aria-hidden="true">→</span>
            </button>
          </div>
        </div>
      </section>
    </Fragment>
  );
}
