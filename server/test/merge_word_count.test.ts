import { describe, expect, it } from "vitest";
import { buildAppendixParagraphs, countDocumentWords, countWords } from "../src/merge/word_count.js";
import { buildDocx } from "./util/docx_builder.js";

describe("merge word count", () => {
  it("counts hyphenated and apostrophized terms consistently", () => {
    expect(countWords("People's long-term costs were 21,000 dollars." )).toBe(7);
  });

  it("excludes a References heading and everything after it", async () => {
    const document = await buildDocx({
      paragraphs: [
        { text: "Five visible words are kept." },
        { text: "References" },
        { text: "These six reference words are not counted." },
      ],
      table: { rows: 1, cols: 2, cellText: "table cell" },
      tableAfter: 0,
    });
    expect(await countDocumentWords(document)).toBe(13);
  });

  it("builds an exact appendix word budget by repeating and truncating the source", async () => {
    const paragraphs = await buildAppendixParagraphs(9000);
    expect(paragraphs.reduce((sum, paragraph) => sum + countWords(paragraph.text), 0)).toBe(9000);
  });
});
