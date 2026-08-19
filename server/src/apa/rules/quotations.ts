import type { ApaRule } from "../types.js";
import { setParagraphIndent } from "../../docx/edit.js";
import { replaceParagraphText, rawParagraphText } from "../../docx/text.js";
import { result, loc, markDocDirty } from "./util.js";
import { excerptOf } from "../types.js";

/** L/M. Direct and block quotation rules. */
export const quotationRules: ApaRule[] = [
  {
    id: "APA-QUOTE-001",
    category: "quotations",
    description: "Quotations of 40+ words are formatted as block quotations.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.longQuoteCandidates.length > 0,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "user_review" | null = null;

      for (const cand of analysis.longQuoteCandidates) {
        checked++;
        const p = model.paragraphs[cand.paragraphIndex]!;
        if (cand.isIndented) {
          passed++;
          continue;
        }
        // Strong-confidence auto-fix: the quotation IS the whole paragraph
        // (aside from a trailing citation) — indent it and drop the marks.
        const raw = rawParagraphText(p.el);
        const wholeMatch = /^\s*[“"]([\s\S]{150,})[”"]\s*(\([^)]*\)\s*\.?\s*)?$/.exec(raw);
        if (fix && wholeMatch) {
          const openChar = raw.trim().charAt(0);
          const closeIdx = raw.lastIndexOf(openChar === "“" ? "”" : '"');
          if (closeIdx > 0) {
            // Remove closing then opening quotation mark.
            replaceParagraphText(p.el, openChar === "“" ? "”" : '"', "");
            replaceParagraphText(p.el, openChar, "");
          }
          setParagraphIndent(model.documentXml, p.el, {
            left: 720,
            firstLine: 0,
          });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-QUOTE-001",
            category: "quotations",
            location: loc(p),
            before: `Quotation of ~${cand.words} words in quotation marks`,
            after: 'Block quotation: 0.5" left indent, quotation marks removed',
            reason:
              "APA 7 formats quotations of 40+ words as indented block quotations without quotation marks.",
            confidence: 0.9,
          });
          continue;
        }
        worst = "user_review";
        ctx.addIssue({
          ruleId: "APA-QUOTE-001",
          category: "quotations",
          severity: "warning",
          status: "user_review",
          message: `A quotation of ~${cand.words} words should be a block quotation.`,
          explanation:
            "The quotation is embedded in a longer paragraph, so it was not converted automatically. Move it to its own paragraph with a 0.5\" indent and no quotation marks, or confirm it is fine as is.",
          location: loc(p),
          originalValue: excerptOf(p.text, 100),
          confidence: 0.7,
          autoFixable: false,
          userResolutionRequired: true,
          resolutionOptions: [
            { id: "acknowledge", label: "I will reformat it", description: "I will convert this to a block quote in Word." },
            { id: "not_a_quote", label: "Not a quotation", description: "This text is not actually a long quotation." },
          ],
        });
      }
      return result("APA-QUOTE-001", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-QUOTE-002",
    category: "quotations",
    description: "Existing block quotations use a 0.5-inch left indent.",
    severity: "warning",
    applies: (ctx) =>
      ctx.analysis.longQuoteCandidates.some((c) => c.isIndented),
    run(ctx, fix) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      for (const cand of analysis.longQuoteCandidates) {
        if (!cand.isIndented) continue;
        checked++;
        const p = model.paragraphs[cand.paragraphIndex]!;
        const left = p.props.leftIndent ?? 0;
        if (Math.abs(left - 720) <= 40) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphIndent(model.documentXml, p.el, { left: 720 });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-QUOTE-002",
            category: "quotations",
            location: loc(p),
            before: `block quote indent ${(left / 1440).toFixed(2)}"`,
            after: 'block quote indent 0.5"',
            reason: "APA 7 block quotations are indented 0.5 inch from the left margin.",
            confidence: 0.85,
          });
        }
      }
      return result("APA-QUOTE-002", checked, passed, fixedCount > 0, null);
    },
  },
];
