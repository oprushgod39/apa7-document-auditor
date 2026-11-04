import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import { Window } from "happy-dom";
import { citationHtmlText, sanitizeCitationHtml, correctedCitationHtml, plainCitationHtml } from "../../web/src/lib/rich_citations.js";
import { correctCitationText } from "../src/apa/citations/correct_text.js";

beforeEach(() => {
  const browser = new Window();
  vi.stubGlobal("document", browser.document);
  vi.stubGlobal("DOMParser", browser.DOMParser);
  vi.stubGlobal("Node", browser.Node);
});
afterEach(() => vi.unstubAllGlobals());

describe("formatted citation paste and copy", () => {
  it("retains italic reference titles while correcting citations in the body", () => {
    const pasted = '<p><b>Research</b> agrees (Smith and Patel 2024).</p><p>References</p><p>Smith, J., &amp; Patel, R. (2024). <span style="font-style:italic">An important book</span>. Publisher.</p>';
    const text = citationHtmlText(pasted);
    const corrected = correctCitationText(text).text;
    const html = correctedCitationHtml(pasted, corrected);
    expect(html).toContain("<em>An important book</em>");
    expect(html).toContain("<strong>Research</strong>");
    expect(citationHtmlText(html)).toBe(corrected);
    expect(corrected).toContain("(Smith & Patel, 2024)");
  });
  it("preserves journal and volume italics across multiple references and repeated edits", () => {
    const pasted = '<div>Evidence (Smith and Patel 2024) and Kumar et al (2023) agrees.</div><div>References</div><div>Smith, J., &amp; Patel, R. (2024). A study. <i>Journal of Research, 12</i>(2), 1–9.</div><div>Kumar, A. (2023). <em>A second book</em>.</div>';
    const html = correctedCitationHtml(pasted, correctCitationText(citationHtmlText(pasted)).text);
    expect(html).toContain("<em>Journal of Research, 12</em>");
    expect(html).toContain("<em>A second book</em>");
    expect(citationHtmlText(html)).toContain("Kumar et al. (2023)");
    expect(correctedCitationHtml(html, citationHtmlText(html))).toBe(html);
  });
  it("removes executable markup and remote content while retaining emphasis", () => {
    const html = sanitizeCitationHtml('<div onclick="alert(1)"><script>alert(2)</script><img src="https://example.com/tracker"><a href="javascript:alert(3)"><span style="font-style:italic;background:url(https://example.com)">Title</span></a><iframe srcdoc="danger"></iframe></div>');
    expect(html).toBe("<div><em>Title</em></div>");
  });
  it("preserves plain-text line breaks and escapes HTML-looking prose", () => {
    const text = "First paragraph.\nReferences\nA title about <AI> & people.";
    const html = plainCitationHtml(text);
    expect(citationHtmlText(html)).toBe(text);
    expect(html).toContain("&lt;AI&gt;");
  });
});
