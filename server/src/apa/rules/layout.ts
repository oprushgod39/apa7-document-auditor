import type { ApaRule } from "../types.js";
import { TWIPS_PER_INCH } from "../types.js";
import {
  setSectionMargins,
  setParagraphSpacing,
  setParagraphRunFonts,
  findStyleEl,
  setStyleRunFormatting,
  setStyleParaFormatting,
  ensurePageNumberHeader,
} from "../../docx/edit.js";
import { result, loc, markDocDirty, markStylesDirty, isReferenceEntry } from "./util.js";

const MARGIN_TOLERANCE = 20; // twips (~0.014")

/** A. Page and document layout rules. */
export const layoutRules: ApaRule[] = [
  {
    id: "APA-LAYOUT-001",
    category: "layout",
    description: "All page margins must be 1 inch.",
    severity: "error",
    applies: () => true,
    run(ctx, fix) {
      const { model } = ctx;
      let checked = 0;
      let passed = 0;
      let fixed = false;
      for (const section of model.sections) {
        checked++;
        const values = [
          section.marginTop,
          section.marginBottom,
          section.marginLeft,
          section.marginRight,
        ];
        const ok = values.every(
          (v) => v != null && Math.abs(v - TWIPS_PER_INCH) <= MARGIN_TOLERANCE
        );
        if (ok) {
          passed++;
          continue;
        }
        const describe = `top ${fmtIn(section.marginTop)}, bottom ${fmtIn(section.marginBottom)}, left ${fmtIn(section.marginLeft)}, right ${fmtIn(section.marginRight)}`;
        if (fix) {
          setSectionMargins(model.documentXml, section.el, {
            top: TWIPS_PER_INCH,
            bottom: TWIPS_PER_INCH,
            left: TWIPS_PER_INCH,
            right: TWIPS_PER_INCH,
          });
          markDocDirty(ctx);
          fixed = true;
          ctx.addChange({
            ruleId: "APA-LAYOUT-001",
            category: "layout",
            location: { description: "Section page setup" },
            before: describe,
            after: 'all margins 1"',
            reason: "APA 7 requires 1-inch margins on all sides.",
            confidence: 0.99,
            documentWide: true,
          });
        } else {
          ctx.addIssue({
            ruleId: "APA-LAYOUT-001",
            category: "layout",
            severity: "error",
            status: "fail",
            message: `Margins are not 1 inch (${describe}).`,
            confidence: 0.99,
            autoFixable: true,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-LAYOUT-001", checked, passed, fixed, fixed || passed === checked ? null : "fail");
    },
  },

  {
    id: "APA-LAYOUT-002",
    category: "layout",
    description: "Document must use an APA-approved font at the approved size.",
    severity: "error",
    applies: () => true,
    async run(ctx, fix) {
      const { model, req } = ctx;
      const stylesDoc = model.stylesXml;
      const targetFont = req.font;
      const targetHalf = req.fontSizePt * 2;
      const approved = new Set(
        req.approvedFonts.map((f) => f.name.toLowerCase())
      );
      // Instructor override narrows the approved set to the mandated font.
      const instructorMandated = ctx.settings.instructor.font != null;

      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;

      for (const p of model.paragraphs) {
        if (p.isEmpty || p.insideTable) continue;
        checked++;
        const font = p.runProps.fontAscii ?? p.runProps.fontHAnsi ?? null;
        const sizeHalf = p.runProps.sizeHalfPoints ?? null;
        const fontOk =
          font != null &&
          (instructorMandated
            ? font.toLowerCase() === targetFont.toLowerCase()
            : approved.has(font.toLowerCase()));
        const approvedEntry = font
          ? req.approvedFonts.find((f) => f.name.toLowerCase() === font.toLowerCase())
          : undefined;
        const expectedHalf = instructorMandated
          ? targetHalf
          : (approvedEntry?.sizePt ?? req.fontSizePt) * 2;
        const sizeOk = sizeHalf != null && fontOk && sizeHalf === expectedHalf;
        if (fontOk && sizeOk) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphRunFonts(model.documentXml, p.el, targetFont, targetHalf);
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-LAYOUT-002",
            category: "layout",
            location: loc(p),
            before: `${font ?? "unspecified font"} ${sizeHalf != null ? sizeHalf / 2 + " pt" : "unspecified size"}`,
            after: `${targetFont} ${req.fontSizePt} pt`,
            reason: "APA 7 requires a consistent approved font throughout.",
            confidence: 0.95,
          });
        } else {
          anyFail = true;
        }
      }

      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-LAYOUT-002",
          category: "layout",
          severity: "error",
          status: "fail",
          message: `${checked - passed} paragraph(s) use a non-approved font or size.`,
          suggestedValue: `${targetFont} ${req.fontSizePt} pt`,
          confidence: 0.95,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }

      // Normalize the Normal style so new/unstyled text inherits correctly.
      if (fix && stylesDoc) {
        const normal = findStyleEl(stylesDoc, "Normal");
        if (normal) {
          setStyleRunFormatting(stylesDoc, normal, {
            font: targetFont,
            halfPoints: targetHalf,
          });
          markStylesDirty(ctx);
        }
      }

      return result("APA-LAYOUT-002", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-LAYOUT-003",
    category: "layout",
    description: "Entire document must be double-spaced.",
    severity: "error",
    applies: () => true,
    run(ctx, fix) {
      const { model } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const p of model.paragraphs) {
        if (p.insideTable) continue; // table cell spacing handled separately
        checked++;
        const line = p.props.line;
        const rule = p.props.lineRule ?? "auto";
        const ok = line === 480 && rule === "auto";
        if (ok) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphSpacing(model.documentXml, p.el, {
            line: 480,
            lineRule: "auto",
          });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-LAYOUT-003",
            category: "layout",
            location: loc(p),
            before: describeSpacing(line, rule),
            after: "double spacing",
            reason: "APA 7 requires double spacing throughout the paper.",
            confidence: 0.98,
          });
        } else {
          anyFail = true;
        }
      }
      // Also normalize Normal style spacing so the default is right.
      if (fix && model.stylesXml) {
        const normal = findStyleEl(model.stylesXml, "Normal");
        if (normal) {
          setStyleParaFormatting(model.stylesXml, normal, {
            line: 480,
            lineRule: "auto",
            before: 0,
            after: 0,
          });
          markStylesDirty(ctx);
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-LAYOUT-003",
          category: "layout",
          severity: "error",
          status: "fail",
          message: `${checked - passed} paragraph(s) are not double-spaced.`,
          confidence: 0.98,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-LAYOUT-003", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-LAYOUT-004",
    category: "layout",
    description: "No extra spacing before or after paragraphs.",
    severity: "warning",
    applies: () => true,
    run(ctx, fix) {
      const { model } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (const p of model.paragraphs) {
        if (p.insideTable) continue;
        checked++;
        const before = p.props.spacingBefore ?? 0;
        const after = p.props.spacingAfter ?? 0;
        if (before === 0 && after === 0) {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphSpacing(model.documentXml, p.el, { before: 0, after: 0 });
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-LAYOUT-004",
            category: "layout",
            location: loc(p),
            before: `spacing before ${before / 20} pt, after ${after / 20} pt`,
            after: "spacing before 0 pt, after 0 pt",
            reason:
              "APA 7 uses double spacing with no extra space between paragraphs.",
            confidence: 0.95,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-LAYOUT-004",
          category: "layout",
          severity: "warning",
          status: "fail",
          message: `${checked - passed} paragraph(s) have extra spacing before/after.`,
          confidence: 0.95,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-LAYOUT-004", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-LAYOUT-005",
    category: "layout",
    description: "Pages must be numbered in the top-right header.",
    severity: "error",
    applies: () => true,
    async run(ctx, fix) {
      const { model, req } = ctx;
      const hasPageField = model.headersFooters.some(
        (h) => h.kind === "header" && h.hasPageField
      );
      if (hasPageField) {
        return result("APA-LAYOUT-005", 1, 1, false, null);
      }
      if (fix) {
        const res = await ensurePageNumberHeader(
          model.pkg,
          model.documentXml,
          model.sections.map((s) => s.el),
          {
            font: req.font,
            halfPoints: req.fontSizePt * 2,
            runningHead: req.runningHeadRequired
              ? ctx.settings.metadata.runningHead ||
                (ctx.settings.metadata.title
                  ? ctx.settings.metadata.title.toUpperCase().slice(0, 50)
                  : undefined)
              : undefined,
          }
        );
        if (res.created) {
          ctx.addChange({
            ruleId: "APA-LAYOUT-005",
            category: "layout",
            location: { description: "Page header" },
            before: "No page numbers",
            after: "Page number field added to header (top right)",
            reason: "APA 7 requires a page number on every page.",
            confidence: 0.97,
            documentWide: true,
          });
          return result("APA-LAYOUT-005", 1, 0, true, null);
        }
      }
      ctx.addIssue({
        ruleId: "APA-LAYOUT-005",
        category: "layout",
        severity: "error",
        status: "fail",
        message: "No page numbers were found in the document header.",
        confidence: 0.9,
        autoFixable: true,
        userResolutionRequired: false,
      });
      return result("APA-LAYOUT-005", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-LAYOUT-006",
    category: "layout",
    description: "Page size should be a standard size (Letter or A4).",
    severity: "warning",
    applies: (ctx) => ctx.model.sections.length > 0,
    run(ctx) {
      let checked = 0;
      let passed = 0;
      for (const s of ctx.model.sections) {
        if (s.pageWidth == null || s.pageHeight == null) continue;
        checked++;
        const letter = close(s.pageWidth, 12240) && close(s.pageHeight, 15840);
        const a4 = close(s.pageWidth, 11906) && close(s.pageHeight, 16838);
        if (letter || a4) passed++;
        else {
          ctx.addIssue({
            ruleId: "APA-LAYOUT-006",
            category: "layout",
            severity: "warning",
            status: "warning",
            message: `Unusual page size (${(s.pageWidth / 1440).toFixed(2)}" × ${(s.pageHeight / 1440).toFixed(2)}"). APA papers normally use Letter or A4.`,
            confidence: 0.9,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-LAYOUT-006", checked, passed, false, passed === checked ? null : "warning");
    },
  },
];

function fmtIn(twips: number | undefined): string {
  return twips == null ? "unset" : `${(twips / TWIPS_PER_INCH).toFixed(2)}"`;
}

function describeSpacing(line: number | undefined, rule: string): string {
  if (line == null) return "unspecified line spacing";
  if (rule === "auto") return `${(line / 240).toFixed(2)}× spacing`;
  return `${line / 20} pt ${rule} spacing`;
}

function close(a: number, b: number, tol = 40): boolean {
  return Math.abs(a - b) <= tol;
}
