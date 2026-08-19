import { describe, expect, it } from "vitest";
import { parseReference } from "../src/apa/references/parser.js";
import { parseCitations } from "../src/apa/citations/parser.js";
import { matchCitationsToReferences } from "../src/apa/references/matcher.js";

describe("reference parser", () => {
  it("parses a journal article with DOI", () => {
    const r = parseReference(
      "Smith, J. K., & Patel, R. (2024). Sleep deprivation and memory consolidation. Journal of Cognitive Neuroscience, 36(2), 112–128. https://doi.org/10.1234/jcn.2024.0112",
      0
    );
    expect(r.surnames).toEqual(["smith", "patel"]);
    expect(r.year).toBe("2024");
    expect(r.doi).toBe("10.1234/jcn.2024.0112");
    expect(r.type).toBe("journal_article");
    expect(r.journal?.volume).toBe("36");
  });

  it("flags old-style doi: labels with a deterministic fix", () => {
    const r = parseReference(
      "Smith, J. (2024). A title. Journal of Things, 1(1), 1–10. doi:10.1234/abc",
      0
    );
    const prob = r.problems.find((x) => x.code === "DOI_OLD_FORMAT");
    expect(prob).toBeDefined();
    expect(prob!.after).toBe("https://doi.org/10.1234/abc");
  });

  it("flags Retrieved from before URLs", () => {
    const r = parseReference(
      "Jones, T. (2022). Web article title. Site Name. Retrieved from https://example.com/article",
      0
    );
    expect(r.problems.some((x) => x.code === "URL_RETRIEVED_FROM")).toBe(true);
  });

  it("classifies books, chapters, videos, dissertations", () => {
    expect(
      parseReference("Adams, B. C. (2021). Foundations of sleep science (2nd ed.). Academic Press.", 0).type
    ).toBe("book");
    expect(
      parseReference(
        "Lee, S. (2020). Chapter title. In A. Editor & B. Editor (Eds.), Big book of things (pp. 1–20). Springer.",
        0
      ).type
    ).toBe("book_chapter");
    expect(
      parseReference("Creator, C. (2023, May 1). Video title [Video]. YouTube. https://youtube.com/watch?v=x", 0).type
    ).toBe("video");
    expect(
      parseReference(
        "Doe, J. (2019). Dissertation title [Doctoral dissertation, Some University]. ProQuest.",
        0
      ).type
    ).toBe("dissertation");
  });

  it("recognizes organization authors", () => {
    const r = parseReference(
      "American Psychological Association. (2020). Publication manual of the American Psychological Association (7th ed.).",
      0
    );
    expect(r.isOrganizationAuthor).toBe(true);
    expect(r.year).toBe("2020");
  });

  it("flags missing year", () => {
    const r = parseReference("Smith, J. Some title without a year. Publisher.", 0);
    expect(r.problems.some((x) => x.code === "MISSING_YEAR")).toBe(true);
  });
});

describe("citation ↔ reference matching", () => {
  const refs = [
    parseReference(
      "Smith, J. K., & Patel, R. (2024). Sleep deprivation and memory. Journal of Cog, 36(2), 112–128.",
      10
    ),
    parseReference(
      "Kumar, A., Lopez, M., & Chen, W. (2023). Sleep architecture. Sleep Research, 12(4), 45–67.",
      11
    ),
    parseReference(
      "Williams, T. (2021). Uncited work. Journal of Nothing, 2(1), 1–5.",
      12
    ),
    parseReference(
      "American Psychological Association. (2020). Publication manual (7th ed.).",
      13
    ),
  ];

  it("matches exact, et al., org-abbreviation; finds missing and uncited", () => {
    const citations = [
      ...parseCitations("Recent work (Smith & Patel, 2024) shows effects.", 1),
      ...parseCitations("Kumar et al. (2023) demonstrated this.", 2),
      ...parseCitations("Style rules exist (APA, 2020).", 3),
      ...parseCitations("A missing one (Johnson, 2023).", 4),
    ];
    const result = matchCitationsToReferences(citations, refs);
    const statuses = result.citationMatches.map((m) => m.status);
    expect(statuses[0]).toBe("exact"); // Smith & Patel 2024
    expect(statuses[1]).toBe("exact"); // Kumar et al. 2023 (3 authors)
    expect(["exact", "probable"]).toContain(statuses[2]); // APA → org
    expect(statuses[3]).toBe("missing_reference"); // Johnson
    const uncited = result.referenceUsage.filter((u) => !u.cited);
    expect(uncited.map((u) => u.referenceIndex)).toEqual([2]); // Williams
  });

  it("distinguishes 2024a from 2024b suffixes", () => {
    const refsAB = [
      parseReference("Smith, J. (2024a). First paper. Journal A, 1(1), 1–10.", 0),
      parseReference("Smith, J. (2024b). Second paper. Journal B, 2(2), 2–20.", 1),
    ];
    const cits = parseCitations("As shown (Smith, 2024b).", 0);
    const result = matchCitationsToReferences(cits, refsAB);
    expect(result.citationMatches[0]!.status).toBe("exact");
    expect(result.citationMatches[0]!.referenceIndexes).toEqual([1]);
  });
});
