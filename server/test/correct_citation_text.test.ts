import { describe, expect, it } from "vitest";
import request from "supertest";
import { correctCitationText } from "../src/apa/citations/correct_text.js";
import { createApp } from "../src/app.js";

describe("plain text citation correction", () => {
  it("applies combined punctuation fixes and preserves prose and line breaks", () => {
    const result = correctCitationText("Results agree (Smith and Patel 2024).\nKumar et al (2023) agrees.");
    expect(result.text).toBe("Results agree (Smith & Patel, 2024).\nKumar et al. (2023) agrees.");
    expect(result.citationsFound).toBe(2);
    expect(correctCitationText(result.text).changes).toHaveLength(0);
  });
  it("handles multiple segments in a single parenthesis", () => {
    expect(correctCitationText("Evidence agrees (Smith and Patel 2024; Kumar et al 2023).").text)
      .toBe("Evidence agrees (Smith & Patel, 2024; Kumar et al., 2023).");
  });
  it("corrects narrative conjunctions and standalone citations", () => {
    expect(correctCitationText("Smith & Patel (2024) agreed.").text).toBe("Smith and Patel (2024) agreed.");
    expect(correctCitationText("Smith 2024").text).toBe("(Smith, 2024)");
    expect(correctCitationText("Smith,2024").text).toBe("(Smith, 2024)");
    expect(correctCitationText("In 2024").text).toBe("In 2024");
  });
  it("abbreviates personal author lists and retains ambiguous lists", () => {
    expect(correctCitationText("As reported (Smith, Jones and Lee, 2024).").text).toBe("As reported (Smith et al., 2024).");
    const input = "Evidence (Smith, Jones and Lee, 2024) differs from (Smith, Evans and Lee, 2024).";
    const result = correctCitationText(input);
    expect(result.text).toContain("Smith, Jones & Lee");
    expect(result.warnings.some((w) => w.includes("disambiguation"))).toBe(true);
  });
  it("preserves group names, quoted text, reference lists and non-citation parentheses", () => {
    const input = "A report (Office of Planning, Research and Evaluation, 2012) found this (n = 40).\nReferences\nSmith, J. (2024). Title. Publisher.";
    expect(correctCitationText(input).text).toBe(input);
    expect(correctCitationText("Plain words with no citations.").text).toBe("Plain words with no citations.");
  });
  it("normalizes comma spacing and flags missing information", () => {
    expect(correctCitationText("Results (Smith,2024).").text).toBe("Results (Smith, 2024).");
    expect(correctCitationText("Results (Smith, 2024.").warnings.some((w) => w.includes("unmatched"))).toBe(true);
    expect(correctCitationText("(Smith)").warnings.some((w) => w.includes("No recognizable"))).toBe(true);
  });
  it("returns corrections without creating a document session and validates text limits", async () => {
    const app = createApp();
    const response = await request(app).post("/api/citations/correct").send({ text: "Results (Smith and Patel 2024)." });
    expect(response.status).toBe(200);
    expect(response.headers["cache-control"]).toBe("no-store");
    expect(response.body.text).toBe("Results (Smith & Patel, 2024).");
    for (const text of [" ", "x".repeat(50_001), 42]) {
      expect((await request(app).post("/api/citations/correct").send({ text })).status).toBe(400);
    }
  });
});
