import { describe, expect, it } from "vitest";
import request from "supertest";
import { referenceItalicSpans } from "../src/apa/references/italic_spans.js";
import { createApp } from "../src/app.js";

export const screenshotReferences = "References\nHillebrand, L., Raisch, S., & Schad, J. (2025). Managing with artificial intelligence: An integrative framework. Academy of Management Annals. https://doi.org/10.5465/annals.2022.0072\nNational Institute of Standards and Technology. (2023). Artificial intelligence risk management framework (AI RMF 1.0). https://doi.org/10.6028/NIST.AI.100-1\nRaisch, S., & Krakowski, S. (2021). Artificial intelligence and management: The automation–augmentation paradox. Academy of Management Review, 46(1), 192–210. https://doi.org/10.5465/amr.2018.0072";

describe("APA reference italic recovery", () => {
  it("identifies precisely the italic spans from the user's three references", () => {
    const result = referenceItalicSpans(screenshotReferences);
    expect(result.spans.map((s) => screenshotReferences.slice(s.start, s.end))).toEqual([
      "Academy of Management Annals", "Artificial intelligence risk management framework (AI RMF 1.0)",
      "Academy of Management Review", "46",
    ]);
    expect(result.warnings).toHaveLength(0);
  });
  it("leaves article titles, issue numbers, page ranges, and URLs outside italic spans", () => {
    const result = referenceItalicSpans(screenshotReferences);
    const selected = result.spans.map((s) => screenshotReferences.slice(s.start, s.end)).join("|");
    expect(selected).not.toContain("Managing with artificial intelligence");
    expect(selected).not.toContain("https:");
    expect(selected).not.toContain("192–210");
    expect(selected).not.toContain("46(1)");
  });
  it("flags unknown types and ignores ordinary body text", () => {
    expect(referenceItalicSpans("Research (Smith, 2024) supports this.").spans).toHaveLength(0);
    const result = referenceItalicSpans("References\nSmith, J. (2024). A chapter. In Editor (Ed.), A collected volume. A publisher.");
    expect(result.spans).toHaveLength(0);
    expect(result.warnings.length).toBeGreaterThan(0);
  });
  it("serves formatting offsets with no stored text and rejects oversized requests", async () => {
    const app = createApp();
    const result = await request(app).post("/api/citations/reference-italics").send({ text: screenshotReferences });
    expect(result.status).toBe(200);
    expect(result.body.spans).toHaveLength(4);
    expect(result.headers["cache-control"]).toBe("no-store");
    expect((await request(app).post("/api/citations/reference-italics").send({ text: "x".repeat(50_001) })).status).toBe(400);
  });
});
