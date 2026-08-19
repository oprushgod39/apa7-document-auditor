import type { ApaRule } from "../types.js";
import {
  ensureParagraphStyle,
  findStyleEl,
  setStyleParaFormatting,
  setStyleRunFormatting,
  setParagraphStyle,
  setParagraphAlignment,
} from "../../docx/edit.js";
import { findHierarchySkips } from "../headings/classifier.js";
import { result, loc, markDocDirty, markStylesDirty } from "./util.js";
import { excerptOf } from "../types.js";

/** D. Heading rules (APA levels 1–5). */

interface HeadingSpec {
  styleId: string;
  name: string;
  alignment: "center" | "left";
  bold: boolean;
  italic: boolean;
  firstLine: number | null;
}

const HEADING_SPECS: Record<number, HeadingSpec> = {
  1: { styleId: "Heading1", name: "heading 1", alignment: "center", bold: true, italic: false, firstLine: null },
  2: { styleId: "Heading2", name: "heading 2", alignment: "left", bold: true, italic: false, firstLine: null },
  3: { styleId: "Heading3", name: "heading 3", alignment: "left", bold: true, italic: true, firstLine: null },
  4: { styleId: "Heading4", name: "heading 4", alignment: "left", bold: true, italic: false, firstLine: 720 },
  5: { styleId: "Heading5", name: "heading 5", alignment: "left", bold: true, italic: true, firstLine: 720 },
};

export const HEADING_RESOLUTION_OPTIONS = [
  { id: "normal", label: "Normal paragraph", description: "This is body text, not a heading." },
  { id: "level1", label: "Level 1 heading", description: "Centered, bold, title case." },
  { id: "level2", label: "Level 2 heading", description: "Flush left, bold." },
  { id: "level3", label: "Level 3 heading", description: "Flush left, bold italic." },
  { id: "level4", label: "Level 4 heading", description: "Indented, bold, inline with text." },
  { id: "level5", label: "Level 5 heading", description: "Indented, bold italic, inline with text." },
];

export const headingRules: ApaRule[] = [
  {
    id: "APA-HEAD-001",
    category: "headings",
    description: "Word heading styles conform to APA levels 1–5 formatting.",
    severity: "warning",
    applies: (ctx) => ctx.model.stylesXml != null && ctx.analysis.headings.length > 0,
    run(ctx, fix) {
      const { model, req, analysis } = ctx;
      const stylesDoc = model.stylesXml!;
      const usedLevels = new Set(
        analysis.headings.filter((h) => h.level > 0).map((h) => h.level)
      );
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;

      for (const level of usedLevels) {
        const spec = HEADING_SPECS[level]!;
        checked++;
        const styleEl =
          findStyleEl(stylesDoc, spec.styleId) ??
          (fix ? ensureParagraphStyle(stylesDoc, spec.styleId, spec.name) : null);
        if (!styleEl) continue;
        if (fix) {
          setStyleRunFormatting(stylesDoc, styleEl, {
            font: req.font,
            halfPoints: req.fontSizePt * 2,
            bold: spec.bold,
            italic: spec.italic,
          });
          setStyleParaFormatting(stylesDoc, styleEl, {
            alignment: spec.alignment,
            line: 480,
            lineRule: "auto",
            before: 0,
            after: 0,
            firstLine: spec.firstLine,
          });
          markStylesDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-HEAD-001",
            category: "headings",
            location: { description: `Style "${spec.styleId}"` },
            before: `Word default heading style`,
            after: `APA Level ${level}: ${spec.alignment}, bold${spec.italic ? " italic" : ""}, ${req.font} ${req.fontSizePt} pt, double-spaced`,
            reason: "APA 7 defines specific formatting for each heading level.",
            confidence: 0.95,
            documentWide: true,
          });
        } else {
          passed++; // check-only mode: reported through audit of paragraphs
        }
      }
      return result("APA-HEAD-001", checked, passed, fixedCount > 0, null);
    },
  },

  {
    id: "APA-HEAD-002",
    category: "headings",
    description:
      "Manually formatted headings should use proper heading levels.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "user_review" | null = null;

      for (const h of analysis.headings) {
        if (h.fromStyle) {
          checked++;
          passed++;
          continue;
        }
        checked++;
        const p = model.paragraphs[h.paragraphIndex]!;
        if (h.confidence === "high" && fix) {
          const spec = HEADING_SPECS[h.level]!;
          if (model.stylesXml) {
            ensureParagraphStyle(model.stylesXml, spec.styleId, spec.name);
            markStylesDirty(ctx);
          }
          setParagraphStyle(model.documentXml, p.el, spec.styleId);
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-HEAD-002",
            category: "headings",
            location: loc(p),
            before: `Manually formatted heading (${h.signals.join(", ")})`,
            after: `Level ${h.level} heading style applied`,
            reason: `Detected as a Level ${h.level} heading with high confidence.`,
            confidence: 0.9,
          });
        } else if (h.confidence === "high" && !fix) {
          ctx.addIssue({
            ruleId: "APA-HEAD-002",
            category: "headings",
            severity: "warning",
            status: "fail",
            message: `"${excerptOf(h.text, 50)}" appears to be a Level ${h.level} heading without a heading style.`,
            location: loc(p),
            confidence: 0.9,
            autoFixable: true,
            userResolutionRequired: false,
          });
        } else {
          worst = "user_review";
          ctx.addIssue({
            ruleId: "APA-HEAD-002",
            category: "headings",
            severity: "warning",
            status: "user_review",
            message: `"${excerptOf(h.text, 50)}" may be a Level ${h.level} heading (${h.confidence} confidence).`,
            explanation: `Signals: ${h.signals.join(", ")}. Confirm the correct level or mark it as a normal paragraph.`,
            location: loc(p),
            originalValue: h.text,
            suggestedValue: `Level ${h.level} heading`,
            confidence: h.confidence === "medium" ? 0.7 : 0.4,
            autoFixable: false,
            userResolutionRequired: true,
            resolutionOptions: HEADING_RESOLUTION_OPTIONS,
          });
        }
      }
      return result("APA-HEAD-002", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-HEAD-003",
    category: "headings",
    description: "Heading hierarchy must not skip levels.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.headings.length > 1,
    run(ctx) {
      const skips = findHierarchySkips(
        ctx.analysis.headings.filter((h) => h.level > 0)
      );
      const checked = ctx.analysis.headings.length;
      for (const s of skips) {
        const p = ctx.model.paragraphs[s.heading.paragraphIndex]!;
        ctx.addIssue({
          ruleId: "APA-HEAD-003",
          category: "headings",
          severity: "warning",
          status: "user_review",
          message: `Heading "${excerptOf(s.heading.text, 40)}" jumps from Level ${s.previousLevel} to Level ${s.heading.level}.`,
          explanation:
            "APA headings should proceed sequentially (Level 1 → 2 → 3). Confirm whether this level is intentional.",
          location: loc(p),
          confidence: 0.7,
          autoFixable: false,
          userResolutionRequired: true,
          resolutionOptions: [
            { id: "intentional", label: "Intentional", description: "The heading level is correct as written." },
            ...HEADING_RESOLUTION_OPTIONS.slice(1),
          ],
        });
      }
      return result(
        "APA-HEAD-003",
        checked,
        checked - skips.length,
        false,
        skips.length > 0 ? "user_review" : null
      );
    },
  },

  {
    id: "APA-HEAD-004",
    category: "headings",
    description: `The paper should not open with an "Introduction" heading.`,
    severity: "info",
    applies: () => true,
    run(ctx) {
      const first = ctx.analysis.headings[0];
      if (first && /^introduction$/i.test(first.text.trim())) {
        const p = ctx.model.paragraphs[first.paragraphIndex]!;
        ctx.addIssue({
          ruleId: "APA-HEAD-004",
          category: "headings",
          severity: "info",
          status: "warning",
          message: `APA 7 papers begin under the paper title without an "Introduction" heading.`,
          explanation:
            "The introduction is understood to be the opening section; the first heading normally comes after it. This is advisory — some instructors request an Introduction heading.",
          location: loc(p),
          confidence: 0.8,
          autoFixable: false,
          userResolutionRequired: false,
        });
        return result("APA-HEAD-004", 1, 0, false, "warning");
      }
      return result("APA-HEAD-004", 1, 1, false, null);
    },
  },
];

export { HEADING_SPECS };
