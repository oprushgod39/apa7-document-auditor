// Utility script: writes the malformed golden fixture to a real file for
// manual testing.  Usage: npx tsx test/util/write_fixture.ts <out.docx>
import fs from "node:fs";
import { buildDocx, malformedStudentPaper } from "./docx_builder.js";

const out = process.argv[2] ?? "Research_Paper.docx";
const buf = await buildDocx(malformedStudentPaper());
fs.writeFileSync(out, buf);
console.log(`wrote ${out} (${buf.length} bytes)`);
