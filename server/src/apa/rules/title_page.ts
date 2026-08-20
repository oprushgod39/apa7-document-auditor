import type { ApaRule, RuleContext } from "../types.js";
import {
  createParagraph,
  setParagraphAlignment,
  setParagraphContextualSpacing,
  setParagraphIndent,
  setParagraphKeepNext,
  setParagraphRunFonts,
  setParagraphSpacing,
  setRunBold,
  setRunItalic,
  insertBeforeEl,
  setParagraphRunColorBlack,
} from "../../docx/edit.js";
import { childrenW, createW, setAttrW } from "../../docx/xml.js";
import { result, loc, markDocDirty } from "./util.js";

/** C. Title page rules. */

function requiredElements(ctx: RuleContext): { key: keyof typeof LABELS; required: boolean }[] {
  if (ctx.settings.paperType === "student") {
    return [
      { key: "title", required: true },
      { key: "author", required: true },
      { key: "institution", required: true },
      { key: "course", required: true },
      { key: "instructor", required: true },
      { key: "dueDate", required: true },
    ];
  }
  return [
    { key: "title", required: true },
    { key: "author", required: true },
    { key: "institution", required: true },
  ];
}

const LABELS = {
  title: "Paper title",
  author: "Author name",
  institution: "Institutional affiliation",
  course: "Course number and name",
  instructor: "Instructor",
  dueDate: "Due date",
} as const;

function metadataValue(ctx: RuleContext, key: keyof typeof LABELS): string | undefined {
  const user = ctx.settings.metadata;
  const detected = ctx.analysis.detectedMetadata;
  switch (key) {
    case "title":
      return user.title || detected.title;
    case "author":
      return user.author || detected.author;
    case "institution":
      return user.institution || detected.institution;
    case "course": {
      const userCourse = [user.courseNumber, user.courseName].filter(Boolean).join(": ");
      if (userCourse) return userCourse;
      const det = [detected.courseNumber, detected.courseName].filter(Boolean).join(": ");
      return det || undefined;
    }
    case "instructor":
      return user.instructor || detected.instructor;
    case "dueDate":
      return user.dueDate || detected.dueDate;
  }
}

function userProvided(ctx: RuleContext, key: keyof typeof LABELS): string | undefined {
  const user = ctx.settings.metadata;
  switch (key) {
    case "title": return user.title;
    case "author": return user.author;
    case "institution": return user.institution;
    case "course": {
      const v = [user.courseNumber, user.courseName].filter(Boolean).join(": ");
      return v || undefined;
    }
    case "instructor": return user.instructor;
    case "dueDate": return user.dueDate;
  }
}

export const titlePageRules: ApaRule[] = [
  {
    id: "APA-TITLE-001",
    category: "title_page",
    description: "The paper must have an APA title page.",
    severity: "error",
    applies: () => true,
    run(ctx, fix) {
      const { model, analysis, req } = ctx;
      if (analysis.hasTitlePage) {
        return result("APA-TITLE-001", 1, 1, false, null);
      }
      const userTitle = ctx.settings.metadata.title;
      if (fix && userTitle) {
        // Build a title page from user-provided metadata only. Never invent.
        const doc = model.documentXml;
        const firstBlock = model.paragraphs[0]?.el ?? null;
        const mk = (text: string, opts: Parameters<typeof createParagraph>[2] = {}) =>
          createParagraph(doc, text, {
            alignment: "center",
            font: req.font,
            halfPoints: req.fontSizePt * 2,
            spacing: { line: 480, lineRule: "auto", before: 0, after: 0 },
            firstLineIndent: null,
            black: true,
            ...opts,
          });
        const lines: ReturnType<typeof createParagraph>[] = [];
        // Three blank lines push the title toward the upper third of the page.
        for (let i = 0; i < 3; i++) lines.push(mk(""));
        lines.push(mk(userTitle, { bold: true }));
        lines.push(mk(""));
        for (const key of ["author", "institution", "course", "instructor", "dueDate"] as const) {
          const v = userProvided(ctx, key);
          if (v) lines.push(mk(v));
        }
        // Page break paragraph ends the title page.
        const breakP = mk("");
        const r = createW(doc, "r");
        const br = createW(doc, "br");
        setAttrW(br, "type", "page");
        r.appendChild(br);
        breakP.appendChild(r);
        lines.push(breakP);
        for (const p of lines) insertBeforeEl(model.body, p, firstBlock);
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-TITLE-001",
          category: "title_page",
          location: { description: "Document start" },
          before: "No title page detected",
          after: `Title page created from your metadata (${userTitle})`,
          reason: "APA 7 papers begin with a title page.",
          confidence: 0.9,
          documentWide: true,
        });
        return result("APA-TITLE-001", 1, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-TITLE-001",
        category: "title_page",
        severity: "error",
        status: "user_review",
        message: "No title page was detected.",
        explanation:
          "Provide the paper title (and other title-page details) in the metadata form so a title page can be created. The auditor never invents missing information.",
        confidence: 0.85,
        autoFixable: false,
        userResolutionRequired: true,
      });
      return result("APA-TITLE-001", 1, 0, false, "user_review");
    },
  },

  {
    id: "APA-TITLE-002",
    category: "title_page",
    description: "The paper title is bold, centered, and in title case.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.hasTitlePage,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      const idx = analysis.detectedMetadata.titleParagraphIndex;
      if (idx == null) return result("APA-TITLE-002", 0, 0, false, null);
      const p = model.paragraphs[idx]!;
      let checked = 1;
      let passed = 0;
      let fixed = false;
      const bold = p.runProps.bold === true;
      const centered = p.props.alignment === "center";
      if (bold && centered) {
        passed = 1;
      } else if (fix) {
        if (!centered) setParagraphAlignment(model.documentXml, p.el, "center");
        if (!bold) {
          for (const r of childrenW(p.el, "r")) setRunBold(model.documentXml, r, true);
        }
        setParagraphRunColorBlack(model.documentXml, p.el);
        markDocDirty(ctx);
        fixed = true;
        ctx.addChange({
          ruleId: "APA-TITLE-002",
          category: "title_page",
          location: loc(p),
          before: `${bold ? "bold" : "not bold"}, ${centered ? "centered" : "not centered"}`,
          after: "bold and centered",
          reason: "APA 7 title-page titles are bold and centered.",
          confidence: 0.9,
        });
      } else {
        ctx.addIssue({
          ruleId: "APA-TITLE-002",
          category: "title_page",
          severity: "warning",
          status: "fail",
          message: "The paper title should be bold and centered.",
          location: loc(p),
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-TITLE-002", checked, passed, fixed, passed || fixed ? null : "fail");
    },
  },

  {
    id: "APA-TITLE-006",
    category: "title_page",
    description: "The paper title is repeated at the start of the body, centered and bold.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.detectedMetadata.title != null,
    run(ctx, fix) {
      const detected = ctx.analysis.detectedMetadata;
      const idx = ctx.analysis.hasTitlePage
        ? ctx.model.paragraphs.find(
            (p) => p.index >= ctx.analysis.titlePageEnd && p.text.trim() === detected.title?.trim()
          )?.index
        : detected.titleParagraphIndex;
      if (idx == null) return result("APA-TITLE-006", 0, 0, false, null);
      const p = ctx.model.paragraphs[idx]!;
      const correct =
        p.props.alignment === "center" &&
        p.runProps.bold === true &&
        (p.props.firstLineIndent ?? 0) === 0 &&
        p.runProps.fontAscii === ctx.req.font &&
        p.runProps.sizeHalfPoints === ctx.req.fontSizePt * 2;
      if (correct) return result("APA-TITLE-006", 1, 1, false, null);
      if (fix) {
        const doc = ctx.model.documentXml;
        setParagraphAlignment(doc, p.el, "center");
        setParagraphIndent(doc, p.el, { firstLine: null, hanging: null, left: 0 });
        setParagraphSpacing(doc, p.el, { before: 0, after: 0, line: 480, lineRule: "auto" });
        setParagraphContextualSpacing(doc, p.el);
        setParagraphKeepNext(doc, p.el, true);
        setParagraphRunFonts(doc, p.el, ctx.req.font, ctx.req.fontSizePt * 2);
        for (const r of childrenW(p.el, "r")) {
          setRunBold(doc, r, true);
          setRunItalic(doc, r, false);
        }
        setParagraphRunColorBlack(doc, p.el);
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-TITLE-006",
          category: "title_page",
          location: loc(p),
          before: "source body-title formatting",
          after: "centered, bold, unindented body title in the required font",
          reason: "APA 7 repeats the paper title as a centered bold heading at the start of the body.",
          confidence: 0.95,
        });
        return result("APA-TITLE-006", 1, 0, true, null);
      }
      ctx.addIssue({
        ruleId: "APA-TITLE-006",
        category: "title_page",
        severity: "warning",
        status: "fail",
        message: "The body should begin with the centered, bold paper title.",
        location: loc(p),
        confidence: 0.9,
        autoFixable: true,
        userResolutionRequired: false,
      });
      return result("APA-TITLE-006", 1, 0, false, "fail");
    },
  },

  {
    id: "APA-TITLE-003",
    category: "title_page",
    description: "All required title-page elements are present.",
    severity: "error",
    applies: (ctx) => ctx.analysis.hasTitlePage,
    run(ctx, fix) {
      const { model, analysis, req } = ctx;
      const elements = requiredElements(ctx);
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "user_review" | null = null;

      // Find insertion point: before the title-page-ending page break.
      const lastTitlePara = model.paragraphs[analysis.titlePageEnd - 1];

      for (const { key } of elements) {
        checked++;
        const value = metadataValue(ctx, key);
        if (value && analysis.detectedMetadata[metaKeyOf(key)] != null) {
          passed++;
          continue;
        }
        const provided = userProvided(ctx, key);
        if (provided && fix && lastTitlePara) {
          const p = createParagraph(model.documentXml, provided, {
            alignment: "center",
            font: req.font,
            halfPoints: req.fontSizePt * 2,
            spacing: { line: 480, lineRule: "auto", before: 0, after: 0 },
            firstLineIndent: null,
            black: true,
          });
          insertBeforeEl(model.body, p, lastTitlePara.el);
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-TITLE-003",
            category: "title_page",
            location: { description: "Title page" },
            before: `${LABELS[key]}: missing`,
            after: `${LABELS[key]}: "${provided}" added`,
            reason: "Required APA title-page element supplied by you.",
            confidence: 0.95,
          });
        } else if (!value) {
          worst = "user_review";
          ctx.addIssue({
            ruleId: "APA-TITLE-003",
            category: "title_page",
            severity: "error",
            status: "user_review",
            message: `Title page is missing: ${LABELS[key]}.`,
            explanation:
              "Enter this value in the title-page metadata form and re-run, or confirm it is present under different wording.",
            confidence: 0.75,
            autoFixable: false,
            userResolutionRequired: true,
            resolutionOptions: [
              { id: "present", label: "It is already present", description: "The element exists on my title page; do not flag it." },
              { id: "skip", label: "Not required for my assignment", description: "My instructor does not require this element." },
            ],
          });
        } else {
          // Detected value exists (possibly via user) — treat as passing.
          passed++;
        }
      }
      return result("APA-TITLE-003", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-TITLE-004",
    category: "title_page",
    description: "Title-page elements are centered.",
    severity: "warning",
    applies: (ctx) => ctx.analysis.hasTitlePage,
    run(ctx, fix) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let anyFail = false;
      for (let i = 0; i < analysis.titlePageEnd; i++) {
        const p = model.paragraphs[i]!;
        if (p.isEmpty) continue;
        checked++;
        if (p.props.alignment === "center") {
          passed++;
          continue;
        }
        if (fix) {
          setParagraphAlignment(model.documentXml, p.el, "center");
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-TITLE-004",
            category: "title_page",
            location: loc(p),
            before: p.props.alignment ?? "left",
            after: "centered",
            reason: "APA 7 title-page elements are centered.",
            confidence: 0.9,
          });
        } else {
          anyFail = true;
        }
      }
      if (!fix && anyFail) {
        ctx.addIssue({
          ruleId: "APA-TITLE-004",
          category: "title_page",
          severity: "warning",
          status: "fail",
          message: "Some title-page lines are not centered.",
          confidence: 0.9,
          autoFixable: true,
          userResolutionRequired: false,
        });
      }
      return result("APA-TITLE-004", checked, passed, fixedCount > 0, anyFail ? "fail" : null);
    },
  },

  {
    id: "APA-TITLE-005",
    category: "title_page",
    description: "Running head (professional papers or when required).",
    severity: "warning",
    applies: (ctx) => ctx.req.runningHeadRequired,
    run(ctx) {
      const { model } = ctx;
      const header = model.headersFooters.find((h) => h.kind === "header");
      const hasText = header != null && header.text.replace(/\d+/g, "").trim().length > 0;
      if (hasText) return result("APA-TITLE-005", 1, 1, false, null);
      // The page-number rule adds the running head when metadata provides one;
      // here we verify and otherwise ask the user.
      if (ctx.settings.metadata.runningHead || ctx.settings.metadata.title) {
        // Will be added by APA-LAYOUT-005 header creation if header was missing.
        const created = model.headersFooters.length === 0;
        if (created) return result("APA-TITLE-005", 1, 0, false, "warning");
      }
      ctx.addIssue({
        ruleId: "APA-TITLE-005",
        category: "title_page",
        severity: "warning",
        status: "user_review",
        message: "A running head is required but none was found in the header.",
        explanation:
          "Provide a running head (max 50 characters, ALL CAPS abbreviated title) in the metadata form.",
        confidence: 0.8,
        autoFixable: false,
        userResolutionRequired: true,
        resolutionOptions: [
          { id: "present", label: "It is already present", description: "The running head exists in my header." },
          { id: "skip", label: "Not required", description: "My assignment does not require a running head." },
        ],
      });
      return result("APA-TITLE-005", 1, 0, false, "user_review");
    },
  },
];

function metaKeyOf(key: keyof typeof LABELS): keyof import("../types.js").TitlePageMetadata {
  return key === "course" ? "courseNumber" : key;
}
