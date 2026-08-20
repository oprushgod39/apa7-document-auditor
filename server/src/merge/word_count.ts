import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DocxPackage } from "../docx/package.js";
import { buildDocumentModel } from "../docx/model.js";
import { descendantsW, paragraphText } from "../docx/xml.js";

const WORD_PATTERN = /[\p{L}\p{N}]+(?:['’.-][\p{L}\p{N}]+)*/gu;
/**
 * Matches a paragraph that is (and is only) a References/Bibliography/Works
 * Cited heading. Exported so the merge pipeline can cut source documents at
 * exactly the same point word counts are computed from — see
 * server/src/merge/html_convert.ts.
 */
export const REFERENCE_HEADING = /^(references|bibliography|works\s+cited)\s*:?[\s.]*$/i;

export function countWords(text: string): number {
  return text.match(WORD_PATTERN)?.length ?? 0;
}

export async function countDocumentWords(buffer: Buffer): Promise<number> {
  const model = await buildDocumentModel(await DocxPackage.load(buffer));
  const reference = model.paragraphs.find((paragraph) => REFERENCE_HEADING.test(paragraph.text.trim()));
  const cutoff = reference?.blockIndex ?? Number.POSITIVE_INFINITY;
  let words = 0;
  for (const paragraph of model.paragraphs) {
    if (paragraph.blockIndex < cutoff) words += countWords(paragraph.text);
  }
  for (const table of model.tables) {
    if (table.blockIndex >= cutoff) continue;
    for (const paragraph of descendantsW(table.el, "p")) {
      words += countWords(paragraphText(paragraph));
    }
  }
  return words;
}

export async function appendixSourceWordCount(): Promise<number> {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const source = await fs.readFile(path.resolve(here, "../../assets/ignore_appendix.txt"), "utf8");
  return countWords(source.replace(/\*\*/g, ""));
}

export interface AppendixParagraph {
  text: string;
  bold: boolean;
}

export async function buildAppendixParagraphs(targetWords: number): Promise<AppendixParagraph[]> {
  if (targetWords <= 0) return [];
  const here = path.dirname(fileURLToPath(import.meta.url));
  const source = await fs.readFile(path.resolve(here, "../../assets/ignore_appendix.txt"), "utf8");
  const lines = source.split(/\r?\n/).map((line) => {
    const bold = line.startsWith("**") && line.endsWith("**") && line.length >= 4;
    return { text: bold ? line.slice(2, -2) : line, bold };
  });
  if (lines.reduce((sum, line) => sum + countWords(line.text), 0) === 0) return [];

  const paragraphs: AppendixParagraph[] = [];
  let remaining = targetWords;
  while (remaining > 0) {
    for (const line of lines) {
      if (remaining <= 0) break;
      const words = [...line.text.matchAll(WORD_PATTERN)];
      if (words.length === 0) {
        paragraphs.push(line);
        continue;
      }
      if (words.length <= remaining) {
        paragraphs.push(line);
        remaining -= words.length;
        continue;
      }
      const last = words[remaining - 1]!;
      paragraphs.push({
        text: line.text.slice(0, (last.index ?? 0) + last[0].length),
        bold: line.bold,
      });
      remaining = 0;
    }
  }
  return paragraphs;
}
