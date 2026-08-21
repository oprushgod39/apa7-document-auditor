import type { ApaRule } from "../types.js";
import { APA_FIRST_LINE_INDENT } from "../types.js";
import {
  setParagraphIndent,
  setParagraphAlignment,
} from "../../docx/edit.js";
import { stripLeadingWhitespace } from "../../docx/text.js";
import {
  result,
  loc,
  markDocDirty,
  isHeadingParagraph,
  isReferenceEntry,
} from "./util.js";

const INDENT_TOLERANCE = 40; // twips

function isDisplayMaterialParagraph(
  p: import("../../docx/model.js").ParagraphModel,
  model: import("../../docx/model.js").DocumentModel,
  analysis: import("../analysis.js").DocumentAnalysis
): boolean {
  const text = p.text.trim();
  if (/^(?:table|figure)\s+\d+\.?$/i.test(text) || /^Note\.\s/i.test(text)) return true;
  if (
    p.index === analysis.bodyStartIndex &&
    analysis.detectedMetadata.title != null &&
    text === analysis.detectedMetadata.title.trim()
  ) return true;
  const previous = [...model.paragraphs]
    .reverse()
    .find((q) => q.index < p.index && !q.isEmpty);
  if (previous != null && /^(?:table|figure)\s+\d+\.?$/i.test(previous.text.trim())) return true;
  // Body text following APA Level 4/5 continues on the same line after a
  // style separator, so it must not receive another first-line indent.
  return previous != null && analysis.headings.some(
    (heading) => heading.paragraphIndex === previous.index && heading.level >= 4
  );
}

/** B. Paragraph rules (body text). */
export const paragraphRules: ApaRule[] = [
  {
    id: "APA-PARA-001",
    category: "paragraph",
    description: "Body paragraphs use a 0.5-inch first-line indent.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis, req } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const p of model.paragraphs) {
        if (
          p.isEmpty ||
          p.insideTable ||
          p.index < analysis.bodyStartIndex ||
          isHeadingParagraph(ctx, p.index) ||
          isReferenceEntry(ctx, p.index) ||
          p.index === analysis.referencesHeadingIndex ||
          p.hasNumbering ||
          p.hasDrawing ||
          isDisplayMaterialParagraph(p, model, analysis)
        ) {
          continue;
        }
        // Block quotes (left-indented long paragraphs) keep their own rules.
        if ((p.props.leftIndent ?? 0) >= 700) continue;
        // Keywords line has its own indent rule.
        if (p.index === analysis.keywordsParagraphIndex) continue;
        // Abstract body is deliberately NOT indented.
        if (analysis.abstractBodyIndexes.includes(p.index)) continue;

        checked++;
        const firstLine = p.props.firstLineIndent ?? 0;
        const hanging = p.props.hangingIndent ?? 0;
        const ok =
          hanging === 0 &&
          Math.abs(firstLine - req.firstLineIndentTwips) <= INDENT_TOLERANCE;
        if (ok) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphIndent(model.documentXml, p.el, {
            firstLine: req.firstLineIndentTwips,
          });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-PARA-001",
            category: "paragraph",
            location: loc(p),
            before:
              hanging > 0
                ? `hanging indent ${(hanging / 1440).toFixed(2)}"`
                : `first-line indent ${(firstLine / 1440).toFixed(2)}"`,
            after: 'first-line indent 0.5"',
            reason: "APA 7 body paragraphs begin with a 0.5-inch indent.",
            confidence: 0.9,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-PARA-001",
          category: "paragraph",
          severity: "warning",
          status: "fail",
          message: `${checked - passed} body paragraph(s) lack a 0.5" first-line indent.`,
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-PARA-001", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-PARA-002",
    category: "paragraph",
    description: "Body text is left-aligned, not justified.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const p of model.paragraphs) {
        if (p.isEmpty || p.insideTable || p.index < analysis.bodyStartIndex) continue;
        if (isHeadingParagraph(ctx, p.index)) continue;
        checked++;
        const align = p.props.alignment ?? "left";
        if (align !== "both" && align !== "distribute") {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphAlignment(model.documentXml, p.el, "left");
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-PARA-002",
            category: "paragraph",
            location: loc(p),
            before: "justified alignment",
            after: "left alignment",
            reason: "APA 7 uses left-aligned (ragged right) text.",
            confidence: 0.95,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-PARA-002",
          category: "paragraph",
          severity: "warning",
          status: "fail",
          message: "Some paragraphs are justified; APA uses left alignment.",
          confidence: 0.95,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-PARA-002", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-PARA-003",
    category: "paragraph",
    description:
      "Manual spaces or tabs must not be used to create the first-line indent.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis, req } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const p of model.paragraphs) {
        if (
          p.isEmpty ||
          p.insideTable ||
          p.index < analysis.bodyStartIndex ||
          isHeadingParagraph(ctx, p.index) ||
          isReferenceEntry(ctx, p.index)
        ) {
          continue;
        }
        const startsWithWhitespace = /^[\s ]{2,}/.test(rawStart(p.text)) || startsWithTab(p);
        if (!startsWithWhitespace) continue;
        checked++;
        if (fix) {
          const removed = stripLeadingWhitespace(p.el);
          if (removed) {
            setParagraphIndent(model.documentXml, p.el, {
              firstLine: req.firstLineIndentTwips,
            });
            markDocDirty(ctx);
            fixedCount++;
            ctx.addChange({
              ruleId: "APA-PARA-003",
              category: "paragraph",
              location: loc(p),
              before: `manual indent (${removed})`,
              after: 'paragraph first-line indent 0.5"',
              reason:
                "Indentation should come from paragraph formatting, not typed spaces/tabs.",
              confidence: 0.9,
            });
          } else {
            passed++;
          }
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-PARA-003",
          category: "paragraph",
          severity: "warning",
          status: "fail",
          message: "Some paragraphs use typed spaces/tabs for indentation.",
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-PARA-003", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-PARA-004",
    category: "paragraph",
    description: "No stray blank paragraphs between body paragraphs.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const paras = model.paragraphs;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (let i = analysis.bodyStartIndex; i < paras.length; i++) {
        const p = paras[i]!;
        if (!p.isEmpty || p.insideTable || p.hasSectPr) continue;
        const prev = paras[i - 1];
        const next = paras[i + 1];
        // Only remove a blank paragraph sandwiched between two text
        // paragraphs — blank paragraphs near tables/figures/breaks may be
        // intentional spacing. Critically, a "blank" paragraph can itself be
        // carrying a manual page break (a run with a plain <w:br/> or, more
        // importantly, <w:br type="page"/> — text content is empty either
        // way, since paragraphText() renders a break as whitespace). Manual
        // page breaks are commonly authored as their own empty paragraph
        // (Word's Ctrl+Enter), so this must never delete `p` itself when it
        // carries one — only checking the neighbors (as this previously did)
        // missed exactly that case and silently destroyed the author's own
        // section-boundary page breaks throughout the document.
        // A page break sitting directly after a table/figure's "Note." line
        // is not a genuine section boundary — some authors add one after
        // every table/figure out of habit — and unlike a real section break
        // (title page → ToC, ToC → Abstract, → References), forcing the
        // next heading onto its own fresh page here is unwanted: content
        // should continue right below the table/figure. This is the one
        // page-break case still safe to strip.
        const breakFollowsNote = prev != null && /^Note\.\s/i.test(prev.text.trim());
        const removable =
          prev != null &&
          next != null &&
          !prev.isEmpty &&
          !next.isEmpty &&
          !prev.hasPageBreakAfterInRuns &&
          !next.props.pageBreakBefore &&
          !prev.hasDrawing &&
          !next.hasDrawing &&
          (!p.hasPageBreakAfterInRuns || breakFollowsNote);
        if (!removable) continue;
        checked++;
        if (fix) {
          p.el.parentNode?.removeChild(p.el);
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-PARA-004",
            category: "paragraph",
            location: { paragraphIndex: p.index, description: "Blank paragraph" },
            before: "Empty paragraph between body paragraphs",
            after: "Removed (double spacing already separates paragraphs)",
            reason: "APA 7 does not use blank lines between paragraphs.",
            confidence: 0.9,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-PARA-004",
          category: "paragraph",
          severity: "warning",
          status: "fail",
          message: "Blank paragraphs found between body paragraphs.",
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-PARA-004", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },
];

function rawStart(text: string): string {
  return text.slice(0, 8);
}

function startsWithTab(p: { el: import("../../docx/xml.js").XElement }): boolean {
  // A leading w:tab inside the first run counts as manual indentation.
  const first = p.el.getElementsByTagNameNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "tab"
  );
  if (first.length === 0) return false;
  // Check if it appears before any text content
  const el = first.item(0)!;
  let node = el.parentNode; // run
  if (!node) return false;
  // crude but effective: tab is in first run and no preceding w:t with text
  const run = node as import("../../docx/xml.js").XElement;
  for (let sib = run.previousSibling; sib; sib = sib.previousSibling) {
    if (sib.nodeType === 1 && (sib as import("../../docx/xml.js").XElement).localName === "r") {
      const t = (sib as import("../../docx/xml.js").XElement).getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "t"
      );
      for (let i = 0; i < t.length; i++) {
        if ((t.item(i)!.textContent ?? "").trim().length > 0) return false;
      }
    }
  }
  for (let sib = el.previousSibling; sib; sib = sib.previousSibling) {
    if (
      sib.nodeType === 1 &&
      (sib as import("../../docx/xml.js").XElement).localName === "t" &&
      ((sib as import("../../docx/xml.js").XElement).textContent ?? "").trim().length > 0
    ) {
      return false;
    }
  }
  return true;
}
