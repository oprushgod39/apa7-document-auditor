// Utility: integrity-check a generated .docx from disk.
// Usage: npx tsx test/util/verify_file.ts <file.docx>
import fs from "node:fs";
import { DocxPackage, verifyDocxIntegrity } from "../../src/docx/package.js";
import { buildDocumentModel } from "../../src/docx/model.js";

const file = process.argv[2];
if (!file) throw new Error("usage: verify_file.ts <file.docx>");
const buf = fs.readFileSync(file);
await verifyDocxIntegrity(buf);
const pkg = await DocxPackage.load(buf);
const model = await buildDocumentModel(pkg);
console.log(
  "INTEGRITY OK —",
  "paragraphs:", model.paragraphs.length,
  "| headers:", model.headersFooters.length,
  "| page-number field:", model.headersFooters.some((h) => h.hasPageField),
  "| sections:", model.sections.length
);
