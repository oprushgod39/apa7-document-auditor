import { describe, expect, it } from "vitest";
import { DocxPackage } from "../src/docx/package.js";
import { buildDocumentModel } from "../src/docx/model.js";
import { classifyHeadings, inferAdjacentHeadingLevels } from "../src/apa/headings/classifier.js";
import { buildDocx } from "./util/docx_builder.js";

const body = "The team works together to understand the different responsibilities and priorities of stakeholders throughout the project and to deliver the intended outcomes for the organisation.";
async function detect(texts: string[]) {
  const model = await buildDocumentModel(await DocxPackage.load(await buildDocx({
    paragraphs: texts.map((text) => ({ text, bold: text !== body, align: "center" as const })),
  })));
  const headings = classifyHeadings(model, { bodyStartIndex: 0, referencesHeadingIndex: null });
  inferAdjacentHeadingLevels(model, headings);
  return headings;
}

describe("adjacent heading hierarchy", () => {
  it("infers a child and matching siblings, then resets at a conventional main section", async () => {
    const headings = await detect([
      "Managing Digital Transformation", "Executive Stakeholder Interests and Priorities",
      body, "Individual Executive Priorities", body, "Conclusion", body,
    ]);
    expect(headings.map((h) => h.level)).toEqual([1, 2, 2, 1]);
    expect(headings[1]!.confidence).toBe("medium");
  });
  it("keeps ordinary separated sections at Level 1", async () => {
    expect((await detect(["Executive Priorities", body, "Project Outcomes", body]))
      .map((h) => h.level)).toEqual([1, 1]);
  });
  it("respects explicitly marked levels", async () => {
    expect((await detect(["Executive Priorities", "[H1] Project Outcomes", body, "[H3] Team Roles", body]))
      .map((h) => h.level)).toEqual([1, 1, 3]);
  });
});
