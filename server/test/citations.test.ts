import { describe, expect, it } from "vitest";
import { parseCitations } from "../src/apa/citations/parser.js";

const p = (text: string) => parseCitations(text, 0);

describe("citation parser", () => {
  it("parses single-author parenthetical", () => {
    const [c] = p("Memory declines under stress (Smith, 2024).");
    expect(c).toMatchObject({ kind: "parenthetical", authors: ["Smith"], year: "2024" });
  });

  it("parses narrative citation", () => {
    const [c] = p("Smith (2024) found strong effects.");
    expect(c).toMatchObject({ kind: "narrative", authors: ["Smith"], year: "2024" });
  });

  it("does not absorb preceding prose into a particle surname citation", () => {
    const [c] = p("The welfare framework in Del Canto et al. (2025) is useful here.");
    expect(c).toMatchObject({ kind: "narrative", authors: ["Del Canto"], etAl: true, year: "2025" });
  });

  it("parses two-author forms with & and and", () => {
    const [a] = p("This was shown before (Smith & Patel, 2024).");
    expect(a!.authors).toEqual(["Smith", "Patel"]);
    const [b] = p("Smith and Patel (2024) demonstrated the effect.");
    expect(b!.authors).toEqual(["Smith", "Patel"]);
    expect(b!.problems).toHaveLength(0);
  });

  it("flags ampersand in narrative and 'and' in parenthetical", () => {
    const [a] = p("Smith & Patel (2024) demonstrated the effect.");
    expect(a!.problems.some((x) => x.code === "AMPERSAND_IN_NARRATIVE")).toBe(true);
    const [b] = p("This was shown (Smith and Patel, 2024).");
    expect(b!.problems.some((x) => x.code === "AND_IN_PARENTHETICAL")).toBe(true);
  });

  it("parses et al. and flags missing period", () => {
    const [a] = p("Earlier work agrees (Kumar et al., 2023).");
    expect(a).toMatchObject({ authors: ["Kumar"], etAl: true, year: "2023" });
    expect(a!.problems).toHaveLength(0);
    const [b] = p("Earlier work agrees (Kumar et al, 2023).");
    expect(b!.problems.some((x) => x.code === "ET_AL_MISSING_PERIOD")).toBe(true);
  });

  it("splits multiple citations in one parenthesis", () => {
    const cits = p("Several studies agree (Smith, 2022; Williams, 2023).");
    expect(cits).toHaveLength(2);
    expect(cits[0]).toMatchObject({ authors: ["Smith"], year: "2022" });
    expect(cits[1]).toMatchObject({ authors: ["Williams"], year: "2023" });
  });

  it("parses page and page-range locators", () => {
    const [a] = p("As noted (Smith, 2024, p. 21).");
    expect(a!.locator).toBe("p. 21");
    const [b] = p("As noted (Smith, 2024, pp. 21–24).");
    expect(b!.locator).toMatch(/pp\.\s*21/);
  });

  it("parses paragraph locators and n.d.", () => {
    const [a] = p("As stated (World Health Organization, n.d., para. 4).");
    expect(a).toMatchObject({ year: "n.d." });
    expect(a!.locator).toMatch(/para/);
  });

  it("parses year suffixes 2024a/2024b", () => {
    const cits = p("Both papers matter (Smith, 2024a; Smith, 2024b).");
    expect(cits.map((c) => c.year)).toEqual(["2024a", "2024b"]);
  });

  it("parses organization authors", () => {
    const [c] = p("Guidelines exist (American Psychological Association, 2020).");
    expect(c!.authors).toEqual(["American Psychological Association"]);
  });

  it("does not treat plain parenthetical years in narrative form twice", () => {
    const cits = p("Smith (2024) argued this.");
    expect(cits).toHaveLength(1);
  });

  it("ignores non-citation parentheticals", () => {
    expect(p("The sample (n = 40) was small.")).toHaveLength(0);
    expect(p("This happened (see Figure 1).")).toHaveLength(0);
  });
});
