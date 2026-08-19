import type { ApaRule } from "../types.js";
import {
  setParagraphAlignment,
  setParagraphIndent,
  setRunBold,
} from "../../docx/edit.js";
import { childrenW } from "../../docx/xml.js";
import { result, loc, markDocDirty } from "./util.js";

/** E. Abstract rules. */
export const abstractRules: ApaRule[] = [
  {
    id: "APA-ABS-001",
    category: "abstract",
    description: `The "Abstract" heading is bold and centered on its own page.`,
    severity: "warning",
    applies: (ctx) => ctx.analysis.abstractHeadingIndex != null,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const p = model.paragraphs[analysis.abstractHeadingIndex!]!;
      const bold = p.runProps.bold === true;
      const centered = p.props.alignment === "center";
      if (bold && centered) return result("APA-ABS-001", 1, 1, false, null);
      if (fix) {
        if (!centered) setParagraphAlignment(model.documentXml, p.el, "center");
        if (!bold) {
          for (const r of childrenW(p.el, "r")) setRunBold(model.documentXml, r, true);
        }
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-ABS-001",
          category: "abstract",
          location: loc(p),
          before: `${bold ? "bold" : "not bold"}, ${centered ? "centered" : "not centered"}`,
          after: "bold and centered",
          reason: `APA 7 formats the "Abstract" label bold and centered.`,
          confidence: 0.95,
        });
        return result("APA-ABS-001", 1, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-ABS-001",
        category: "abstract",
        severity: "warning",
        status: "fail",
        message: `The "Abstract" heading should be bold and centered.`,
        location: loc(p),
        confidence: 0.95,
        autoFixable: true,
        userResolutionRequired: false,
      });
      return result("APA-ABS-001", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-ABS-002",
    category: "abstract",
    description: "The first abstract paragraph is not indented.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.abstractBodyIndexes.length > 0,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const first = model.paragraphs[analysis.abstractBodyIndexes[0]!]!;
      const indent = first.props.firstLineIndent ?? 0;
      if (indent === 0) return result("APA-ABS-002", 1, 1, false, null);
      if (fix) {
        setParagraphIndent(model.documentXml, first.el, { firstLine: 0 });
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-ABS-002",
          category: "abstract",
          location: loc(first),
          before: `first-line indent ${(indent / 1440).toFixed(2)}"`,
          after: "no first-line indent",
          reason: "APA 7 abstracts are written as a single unindented block.",
          confidence: 0.9,
        });
        return result("APA-ABS-002", 1, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-ABS-002",
        category: "abstract",
        severity: "warning",
        status: "fail",
        message: "The abstract paragraph should not be indented.",
        location: loc(first),
        confidence: 0.9,
        autoFixable: true,
        userResolutionRequired: false,
      });
      return result("APA-ABS-002", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-ABS-003",
    category: "abstract",
    description: "Abstract presence matches assignment requirements.",
    severity: "warning",
    applies: (ctx) => ctx.req.abstractRequired != null,
    run(ctx) {
      const required = ctx.req.abstractRequired!;
      const present = ctx.analysis.abstractHeadingIndex != null;
      if (required === present) return result("APA-ABS-003", 1, 1, false, null);
      ctx.addIssue({
        ruleId: "APA-ABS-003",
        category: "abstract",
        severity: "warning",
        status: "user_review",
        message: required
          ? "Your assignment requires an abstract, but none was detected."
          : "Your assignment says no abstract, but an Abstract section was detected.",
        explanation: required
          ? "The auditor does not write abstracts. Add one to the document and re-run."
          : "Confirm whether the abstract should be removed (the auditor will not delete your text automatically).",
        confidence: 0.85,
        autoFixable: false,
        userResolutionRequired: true,
        resolutionOptions: [
          { id: "acknowledge", label: "Acknowledged", description: "I will handle this in the document myself." },
          { id: "skip", label: "Requirement does not apply", description: "This requirement was misread; ignore it." },
        ],
      });
      return result("APA-ABS-003", 1, 0, false, "user_review");
    },
  },

  {
    id: "APA-ABS-004",
    category: "abstract",
    description: `The keywords line is indented and begins with italic "Keywords:".`,
    severity: "info",
    applies: (ctx) => ctx.analysis.keywordsParagraphIndex != null,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const p = model.paragraphs[analysis.keywordsParagraphIndex!]!;
      const indent = p.props.firstLineIndent ?? 0;
      const indentOk = Math.abs(indent - 720) <= 40;
      const firstRun = p.runs.find((r) => r.text.trim().length > 0);
      const italicLabel =
        firstRun != null &&
        /^keywords?:?/i.test(firstRun.text.trim()) &&
        firstRun.effective.italic === true;
      if (indentOk && italicLabel) return result("APA-ABS-004", 1, 1, false, null);
      if (fix && !indentOk) {
        setParagraphIndent(model.documentXml, p.el, { firstLine: 720 });
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-ABS-004",
          category: "abstract",
          location: loc(p),
          before: `keywords indent ${(indent / 1440).toFixed(2)}"`,
          after: 'keywords indent 0.5"',
          reason: "APA 7 indents the keywords line like a paragraph.",
          confidence: 0.9,
        });
      }
      if (!italicLabel) {
        ctx.addIssue({
          ruleId: "APA-ABS-004",
          category: "abstract",
          severity: "info",
          status: "warning",
          message: `The word "Keywords:" should be italicized.`,
          explanation:
            "The italic label could not be verified or applied automatically because the label and keywords share a run. Italicize “Keywords:” (label only) in Word.",
          location: loc(p),
          confidence: 0.7,
          autoFixable: false,
          userResolutionRequired: false,
        });
        return result("APA-ABS-004", 1, 0, fix && !indentOk, "warning");
      }
      return result("APA-ABS-004", 1, 0, true, null);
    },
  },
];
