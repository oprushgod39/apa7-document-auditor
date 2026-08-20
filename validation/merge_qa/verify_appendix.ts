import fs from "node:fs/promises";
import { DocxPackage } from "../../server/src/docx/package.js";
import { buildDocumentModel } from "../../server/src/docx/model.js";
import { countWords } from "../../server/src/merge/word_count.js";

async function main() {
  const buffer = await fs.readFile("validation/merge_qa/merged-9000-v3.docx");
  const model = await buildDocumentModel(await DocxPackage.load(buffer));
  const ignore = model.paragraphs.findIndex((paragraph) => paragraph.text.trim() === "IGNORE");
  const appendixWords = model.paragraphs.slice(ignore + 1).reduce((sum, paragraph) => sum + countWords(paragraph.text), 0);
  console.log({ appendixWords, ignore, tables: model.tables.length, images: model.imageCount });
}

void main();
