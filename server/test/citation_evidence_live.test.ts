import { describe, it, expect } from "vitest";
import { parseReference } from "../src/apa/references/parser.js";
import { retrieveSource } from "../src/apa/citations/evidence/sources.js";
import { assessParagraph } from "../src/apa/citations/evidence/openai.js";
import { planEdits } from "../src/apa/citations/evidence/edits.js";

// Opt-in: real network retrieval and paid model inference. Never counted as
// passed semantic validation when credentials or source access are absent.
describe.skipIf(process.env.RUN_LIVE_CITATION_EVAL !== "1")("live Mississippi attribution regression", () => {
  it("grounds each distinct research claim and rejects the unsupported $1.20 assertion", async () => {
    expect(process.env.OPENAI_API_KEY).toBeTruthy(); expect(process.env.OPENAI_CITATION_MODEL).toBeTruthy();
    const refs = [
      "Barr, A., & Gibbs, C. R. (2022). Breaking the cycle? Intergenerational effects of an antipoverty program in early childhood. Journal of Political Economy, 130(12), 3253–3285. https://doi.org/10.1086/720764",
      "Dahl, G. B., & Lochner, L. (2012). The impact of family income on child achievement: Evidence from the earned income tax credit. American Economic Review, 102(5), 1927–1956. https://doi.org/10.1257/aer.102.5.1927",
      "Carneiro, P., & Ginja, R. (2014). Long-term impacts of compensatory preschool on health and behavior: Evidence from Head Start. American Economic Journal: Economic Policy, 6(4), 135–173. https://doi.org/10.1257/pol.6.4.135",
      "Sowell, T. (2007, August 28). No trade-offs? RealClearPolitics. https://www.realclearpolitics.com/articles/2007/08/no_tradeoffs.html",
    ].map(parseReference);
    const sources = await Promise.all(refs.map(retrieveSource));
    // No inaccessible-source result can be used to claim this live eval passed.
    expect(sources.every((s) => s.status === "available"), JSON.stringify(sources.map((s) => ({ id: s.id, status: s.status, scope: s.scope })))).toBe(true);
    const sentences = ["Early childhood antipoverty programs can have intergenerational effects.", "Income changes induced by the earned income tax credit affect children's achievement.", "Head Start can affect later health and behavior."];
    const text = sentences.join(" ") + " A dollar of family income is worth exactly $1.20 of achieving reading and math standards.";
    const first = await assessParagraph(text, sources, "", process.env.OPENAI_CITATION_MODEL!);
    const decisions = await assessParagraph(text, sources, "", process.env.OPENAI_CITATION_MODEL!, first);
    for (const [index, sentence] of sentences.entries()) expect(decisions.some((d) => d.sentence === sentence && d.status === "supported" && d.sourceIds.includes(index))).toBe(true);
    const unsupported = decisions.filter((d) => d.claim.includes("1.20"));
    expect(unsupported.length).toBeGreaterThan(0);
    expect(planEdits(text, 0, unsupported, sources).edits).toHaveLength(0);
  }, 180000);
});
