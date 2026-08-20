// @ts-expect-error pdfjs-dist's legacy Node build ships no types for this
// deep import path; the public API surface used below matches its .d.ts.
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.js";

/**
 * Extracts plain text (and page count) from a PDF buffer for assertions in
 * tests.
 *
 * Note: this project's merge tests previously used `pdf-parse`, but its
 * bundled pdf.js builds (v1.10.100 / v1.10.88 / v2.0.550) fail to parse
 * pdfkit-generated PDFs at all ("bad XRef entry"), even though the files
 * are well-formed (verified byte-for-byte against the PDF xref spec and
 * independently confirmed openable by this same pdfjs-dist version). That
 * is a bug/incompatibility in pdf-parse's old vendored parsers, not in the
 * generated PDFs. `pdfjs-dist` is already a real dependency of this repo
 * (the web app uses it to render PDFs), actively maintained, and parses
 * pdfkit's output correctly — so tests use it directly instead.
 */
export async function extractPdfText(buffer: Buffer): Promise<{ text: string; numPages: number }> {
  const data = new Uint8Array(buffer);
  const loadingTask = (pdfjsLib as any).getDocument({
    data,
    disableWorker: true,
    isEvalSupported: false,
    useSystemFonts: true,
  });
  const doc = await loadingTask.promise;
  let text = "";
  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();
    text += content.items.map((item: any) => item.str ?? "").join(" ") + " ";
  }
  return { text, numPages: doc.numPages };
}
