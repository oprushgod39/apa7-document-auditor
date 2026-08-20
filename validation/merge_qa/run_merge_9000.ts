import fs from "node:fs/promises";
import { mergeDocuments } from "../../server/src/merge/merge.js";

async function main() {
  const [alpha, beta] = await Promise.all([
    fs.readFile("validation/merge_qa/alpha.docx"),
    fs.readFile("validation/merge_qa/beta.docx"),
  ]);
  const output = await mergeDocuments([
    { name: "Alpha Student", originalName: "alpha.docx", buffer: alpha },
    { name: "Beta Student", originalName: "beta.docx", buffer: beta },
  ], 9000);
  await fs.writeFile("validation/merge_qa/merged-9000-v3.docx", output);
}

void main();
