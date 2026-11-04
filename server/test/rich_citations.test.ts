import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import { Window } from "happy-dom";
import { citationHtmlText, sanitizeCitationHtml, correctedCitationHtml, plainCitationHtml, applyReferenceItalics } from "../../web/src/lib/rich_citations.js";
import { referenceItalicSpans } from "../src/apa/references/italic_spans.js";
import { correctCitationText } from "../src/apa/citations/correct_text.js";

beforeEach(() => {
  const browser = new Window();
  vi.stubGlobal("document", browser.document);
  vi.stubGlobal("DOMParser", browser.DOMParser);
  vi.stubGlobal("Node", browser.Node);
});
afterEach(() => vi.unstubAllGlobals());

describe("formatted citation paste and copy", () => {
  it("recovers APA italics in a plain-text copy without changing any characters", () => {
    const text = "References\nHillebrand, L., Raisch, S., & Schad, J. (2025). Managing with artificial intelligence: An integrative framework. Academy of Management Annals. https://doi.org/10.5465/annals.2022.0072\nNational Institute of Standards and Technology. (2023). Artificial intelligence risk management framework (AI RMF 1.0). https://doi.org/10.6028/NIST.AI.100-1\nRaisch, S., & Krakowski, S. (2021). Artificial intelligence and management: The automation–augmentation paradox. Academy of Management Review, 46(1), 192–210. https://doi.org/10.5465/amr.2018.0072";
    const spans = referenceItalicSpans(text).spans;
    const html = applyReferenceItalics(plainCitationHtml(text), spans);
    expect(html).toContain("<em>Academy of Management Annals</em>");
    expect(html).toContain("<em>Artificial intelligence risk management framework (AI RMF 1.0)</em>");
    expect(html).toContain("<em>Academy of Management Review</em>, <em>46</em>(1)");
    expect(citationHtmlText(html)).toBe(text);
    expect(applyReferenceItalics(html, spans)).toBe(html);
  });
  it("recognizes emphasis carried by a semantic italic class without a stylesheet", () => {
    expect(sanitizeCitationHtml('<span class="italic">Journal title</span>')).toBe("<em>Journal title</em>");
  });
  it("recognizes Word class-based and oblique emphasis without retaining the stylesheet", () => {
    const pasted = '<html><head><style>span.BookTitle {font-style:italic; color:red} .Journal, .Volume {font:italic 12pt serif; background:url(https://example.com)}</style></head><body><p>References</p><p>Author (2024). <span class="BookTitle">Book title</span>.</p><p><span class="Journal">Journal title</span>, <span class="Volume">12</span>.</p><p><span style="font-style:oblique">Another title</span>.</p></body></html>';
    const html = sanitizeCitationHtml(pasted);
    expect(html).toContain("<em>Book title</em>");
    expect(html).toContain("<em>Journal title</em>");
    expect(html).toContain("<em>12</em>");
    expect(html).toContain("<em>Another title</em>");
    expect(html).not.toContain("<style");
    expect(html).not.toContain("example.com");
    expect(correctedCitationHtml(html, citationHtmlText(html))).toBe(html);
  });
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
