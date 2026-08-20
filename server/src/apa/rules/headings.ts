import type { ApaRule } from "../types.js";
import {
  ensureParagraphStyle,
  findStyleEl,
  setStyleParaFormatting,
  setStyleRunFormatting,
  setParagraphStyle,
  setParagraphAlignment,
  setParagraphRunColorBlack,
  setParagraphRunFonts,
  setParagraphIndent,
  setParagraphSpacing,
  setParagraphContextualSpacing,
  setParagraphStyleSeparator,
  setRunBold,
  setRunItalic,
  setRunColorBlack,
  setRunUnderlineNone,
} from "../../docx/edit.js";
import { childrenW } from "../../docx/xml.js";
import { replaceParagraphText } from "../../docx/text.js";
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

const TITLE_CASE_MINOR_WORDS = new Set([
  "a", "an", "and", "as", "at", "but", "by", "for", "in", "nor", "of",
  "on", "or", "per", "the", "to", "up", "via", "with", "yet",
]);
const HEADING_ACRONYMS: Record<string, string> = {
  ai: "AI", apa: "APA", cpi: "CPI", eu: "EU", gdp: "GDP", uk: "UK",
  us: "US", usa: "USA",
};

/** Conservative APA-style title case for text supplied after a heading marker. */
function toHeadingTitleCase(text: string): string {
  const words = text.split(/(\s+)/);
  const lexical = words.filter((word) => !/^\s+$/.test(word));
  let position = 0;
  return words.map((word) => {
    if (/^\s+$/.test(word)) return word;
    const current = position++;
    const match = /^(\W*)([\p{L}\p{N}.''’-]+)(\W*)$/u.exec(word);
    if (!match) return word;
    const [, leading, core, trailing] = match;
    // Preserve deliberate acronyms and mixed-case names such as U.S. or GenAI.
    if (/[A-Z].*[A-Z]/.test(core!) || /[a-z][A-Z]/.test(core!)) return word;
    const lower = core!.toLocaleLowerCase("en-US");
    const acronym = HEADING_ACRONYMS[lower];
    const isEdge = current === 0 || current === lexical.length - 1;
    const converted = acronym ??
      (!isEdge && TITLE_CASE_MINOR_WORDS.has(lower)
        ? lower
        : lower.replace(/^\p{L}/u, (letter) => letter.toLocaleUpperCase("en-US")));
    return `${leading}${converted}${trailing}`;
  }).join("");
}

function applyHeadingDirectFormatting(
  ctx: Parameters<ApaRule["run"]>[0],
  paragraphIndex: number,
  level: number
): void {
  const p = ctx.model.paragraphs[paragraphIndex]!;
  const spec = HEADING_SPECS[level]!;
  const doc = ctx.model.documentXml;
  setParagraphStyle(doc, p.el, spec.styleId);
  setParagraphAlignment(doc, p.el, spec.alignment);
  setParagraphIndent(doc, p.el, { firstLine: spec.firstLine, hanging: null });
  setParagraphSpacing(doc, p.el, { before: 0, after: 0, line: 480, lineRule: "auto" });
  setParagraphContextualSpacing(doc, p.el);
  setParagraphRunFonts(doc, p.el, ctx.req.font, ctx.req.fontSizePt * 2);
  for (const r of childrenW(p.el, "r")) {
    setRunBold(doc, r, spec.bold);
    setRunItalic(doc, r, spec.italic);
    setRunColorBlack(doc, r);
    setRunUnderlineNone(doc, r);
  }
  setParagraphRunColorBlack(doc, p.el);
  const runIn = level >= 4;
  setParagraphStyleSeparator(doc, p.el, runIn);
  if (runIn) {
    const next = ctx.model.paragraphs.find(
      (candidate) => candidate.index > paragraphIndex && !candidate.isEmpty
    );
    if (next) {
      // The style separator makes this paragraph continue after the heading.
      // Remove its ordinary first-line indent so it does not add a second tab.
      setParagraphIndent(doc, next.el, { firstLine: null, hanging: null, left: 0 });
    }
  }
  markDocDirty(ctx);
}

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
            black: true,
          });
          setStyleParaFormatting(stylesDoc, styleEl, {
            alignment: spec.alignment,
            line: 480,
            lineRule: "auto",
            before: 0,
            after: 0,
            firstLine: spec.firstLine,
            contextualSpacing: true,
            // Levels 4 and 5 are run-in headings. Keeping their hidden
            // paragraph mark with the next paragraph defeats Word's style
            // separator and leaves the body on a new physical line.
            keepNext: level < 4,
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
      // Paragraph properties and direct run formatting override Word styles.
      // Normalize both layers so a left-aligned source Heading 1 or a direct
      // Aptos run cannot defeat the APA style definition.
      if (fix) {
        for (const heading of analysis.headings) {
          applyHeadingDirectFormatting(ctx, heading.paragraphIndex, heading.level);
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
          const sourceHeadingText = h.marker
            ? toHeadingTitleCase(h.marker.cleanText)
            : h.text;
          const finalHeadingText =
            h.level >= 4 && !/[.!?]$/.test(sourceHeadingText)
              ? `${sourceHeadingText}.`
              : sourceHeadingText;
          if (h.marker || h.level >= 4) {
            // A hidden style-separator paragraph mark has zero rendered width.
            // Preserve one literal space after run-in headings so Word displays
            // "Heading. Body" instead of "Heading.Body".
            replaceParagraphText(
              p.el,
              p.text.trim(),
              h.level >= 4 ? `${finalHeadingText} ` : finalHeadingText
            );
          }
          applyHeadingDirectFormatting(ctx, h.paragraphIndex, h.level);
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-HEAD-002",
            category: "headings",
            location: loc(p),
            before: h.marker
              ? p.text.trim()
              : `Manually formatted heading (${h.signals.join(", ")})`,
            after: h.marker
              ? finalHeadingText
              : `Level ${h.level} heading style applied`,
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
