import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { AppError } from "../errors.js";
import { buildAppendixParagraphs } from "./word_count.js";

const execFileAsync = promisify(execFile);

export interface MergeInput {
  name: string;
  originalName: string;
  buffer: Buffer;
}

export async function mergeDocuments(inputs: MergeInput[], appendixWords: number): Promise<Buffer> {
  if (process.platform === "win32") {
    const { stdout } = await execFileAsync(
      "tasklist.exe",
      ["/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
      { windowsHide: true }
    );
    if (/\bWINWORD\.EXE\b/i.test(stdout)) {
      throw new AppError(
        "PROCESSING_FAILED",
        "Save and close Microsoft Word before merging so your open documents stay untouched, then try again.",
        409
      );
    }
  }
  const here = path.dirname(fileURLToPath(import.meta.url));
  const serverRoot = path.resolve(here, "../..");
  const scriptPath = path.join(serverRoot, "scripts", "merge_documents.ps1");
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "apa7-merge-"));
  const outputPath = path.join(tempDir, "merged.docx");
  const manifestPath = path.join(tempDir, "manifest.json");

  try {
    const documents = [];
    for (let index = 0; index < inputs.length; index++) {
      const item = inputs[index]!;
      const sourcePath = path.join(tempDir, `source-${String(index + 1).padStart(2, "0")}.docx`);
      await fs.writeFile(sourcePath, item.buffer);
      documents.push({ path: sourcePath, name: item.name });
    }
    await fs.writeFile(
      manifestPath,
      JSON.stringify({ documents, appendixWords, appendixParagraphs: await buildAppendixParagraphs(appendixWords), outputPath }),
      "utf8"
    );

    await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-ManifestPath", manifestPath],
      { timeout: 180_000, windowsHide: true, maxBuffer: 2 * 1024 * 1024 }
    );
    return await fs.readFile(outputPath);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/ActiveX|COM|Word\.Application|class not registered/i.test(message)) {
      throw new AppError("PROCESSING_FAILED", "Microsoft Word is required on this computer to preserve every table, image, and style during merging.", 503);
    }
    throw new AppError("PROCESSING_FAILED", "The documents could not be merged. Close any Word dialogs and try again.", 500);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}
