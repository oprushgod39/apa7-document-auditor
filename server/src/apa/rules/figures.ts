import type { ApaRule } from "../types.js";
import {
  replaceParagraphRuns,
  setParagraphAlignment,
  setParagraphIndent,
  setParagraphKeepLines,
  setParagraphKeepNext,
  setParagraphRunFonts,
  setParagraphSpacing,
} from "../../docx/edit.js";
import { result, markDocDirty } from "./util.js";
import { formatCaptionParagraph } from "./captions.js";

/** O. Figure rules — audit only plus caption formatting. Images are never touched. */
export const figureRules: ApaRule[] = [
  {
    id: "APA-FIGURE-001",
    category: "figures",
    description: `Each figure has a bold "Figure N" label and an italic title.`,
    severity: "warning",
    applies: (ctx) => ctx.model.imageCount > 0,
    run(ctx, fix) {
      const { model } = ctx;
      const figureParas = model.paragraphs.filter((p) => p.hasDrawing && !p.insideTable);
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "user_review" | null = null;

      figureParas.forEach((figPara, fi) => {
        checked++;
        // Label may be above (APA 7) — check the two non-empty paragraphs before.
        const before = model.paragraphs
          .filter((p) => p.index < figPara.index && !p.isEmpty && !p.hasDrawing)
          .slice(-2);
        const after = model.paragraphs
          .filter((p) => p.index > figPara.index && !p.isEmpty && !p.hasDrawing)
          .slice(0, 2);
        const labelRe = /^figure\s+\d+\.?$/i;
        const numberPara =
          before.find((p) => labelRe.test(p.text.trim())) ??
          after.find((p) => labelRe.test(p.text.trim()));
        if (!numberPara) {
          worst = "user_review";
          ctx.addIssue({
            ruleId: "APA-FIGURE-001",
            category: "figures",
            severity: "warning",
            status: "user_review",
            message: `An image (figure ${fi + 1}) has no "Figure N" label.`,
            explanation:
              'APA 7 places a bold "Figure N" label and an italic title above each figure. The auditor does not invent titles — add them, or confirm the image is decorative.',
            location: { paragraphIndex: figPara.index, description: "Embedded image" },
            confidence: 0.8,
            autoFixable: false,
            userResolutionRequired: true,
            resolutionOptions: [
              { id: "will_add", label: "I will add the label", description: "I will add the Figure N label and title." },
              { id: "decorative", label: "Decorative image", description: "This image is not an APA figure." },
            ],
          });
          return;
        }
        const numberBeforeFigure = numberPara.index < figPara.index;
        const titlePara = numberBeforeFigure
          ? before.find((p) => p.index > numberPara.index && p.index < figPara.index)
          : after.find((p) => p.index > numberPara.index);
        const ok =
          numberPara.runProps.bold === true &&
          numberPara.runProps.italic !== true &&
          numberPara.props.alignment === "left" &&
          titlePara != null &&
          titlePara.runProps.italic === true &&
          titlePara.runProps.bold !== true &&
          titlePara.props.alignment === "left";
        if (fix) {
          formatCaptionParagraph(ctx, numberPara, "label");
          if (titlePara) formatCaptionParagraph(ctx, titlePara, "title");
          const doc = model.documentXml;
          setParagraphAlignment(doc, figPara.el, "center");
          setParagraphIndent(doc, figPara.el, { firstLine: null, hanging: null, left: 0 });
          setParagraphSpacing(doc, figPara.el, { before: 0, after: 0, line: 240, lineRule: "auto" });
          setParagraphKeepLines(doc, figPara.el, true);
          const notePara = model.paragraphs.find(
            (p) => p.index > figPara.index && !p.isEmpty && !p.hasDrawing
          );
          const noteMatch = notePara ? /^(Note\.\s*)/i.exec(notePara.text) : null;
          setParagraphKeepNext(doc, figPara.el, noteMatch != null);
          if (notePara && noteMatch) {
            setParagraphAlignment(doc, notePara.el, "left");
            setParagraphIndent(doc, notePara.el, { firstLine: null, hanging: null, left: 0 });
            setParagraphSpacing(doc, notePara.el, { before: 0, after: 0, line: 240, lineRule: "auto" });
            setParagraphKeepLines(doc, notePara.el, true);
            setParagraphRunFonts(doc, notePara.el, ctx.req.font, ctx.req.fontSizePt * 2);
            replaceParagraphRuns(doc, notePara.el, [
              { text: noteMatch[1]!, italic: true, bold: false, font: ctx.req.font, halfPoints: ctx.req.fontSizePt * 2, black: true },
              { text: notePara.text.slice(noteMatch[1]!.length), italic: false, bold: false, font: ctx.req.font, halfPoints: ctx.req.fontSizePt * 2, black: true },
            ]);
          }
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-FIGURE-001",
            category: "figures",
            location: { paragraphIndex: numberPara.index, excerpt: numberPara.text.trim() },
            before: "source figure label/title formatting",
            after: "left-aligned bold figure label and left-aligned italic non-bold title",
            reason: "APA 7 figure numbers and titles use distinct formatting on separate lines.",
            confidence: 0.95,
          });
          if (!titlePara) {
            worst = "user_review";
            ctx.addIssue({
              ruleId: "APA-FIGURE-001",
              category: "figures",
              severity: "warning",
              status: "user_review",
              message: `Figure ${fi + 1} has a label but no separate title/explanation line.`,
              explanation: "Add a descriptive title immediately below the Figure N label. The auditor does not invent it.",
              location: { paragraphIndex: figPara.index, description: "Embedded image" },
              confidence: 0.8,
              autoFixable: false,
              userResolutionRequired: true,
            });
          }
        } else if (ok) {
          passed++;
        } else {
          ctx.addIssue({
            ruleId: "APA-FIGURE-001",
            category: "figures",
            severity: "warning",
            status: "fail",
            message: `Figure ${fi + 1} label/title formatting does not follow APA (left-aligned bold number; left-aligned italic non-bold title).`,
            location: { paragraphIndex: figPara.index, description: "Embedded image" },
            confidence: 0.9,
            autoFixable: titlePara != null,
            userResolutionRequired: titlePara == null,
          });
        }
      });
      return result("APA-FIGURE-001", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-FIGURE-002",
    category: "figures",
    description: "Each labeled figure is referred to in the text.",
    severity: "info",
    applies: (ctx) => ctx.model.imageCount > 0,
    run(ctx) {
      const { model, analysis } = ctx;
      const labels = model.paragraphs
        .filter((p) => /^figure\s+\d+\.?$/i.test(p.text.trim()))
        .map((p) => Number.parseInt(/(\d+)/.exec(p.text)![1]!, 10));
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      const bodyText = model.paragraphs
        .filter((p) => !p.insideTable && p.index >= analysis.bodyStartIndex)
        .map((p) => p.text)
        .join("\n");
      for (const n of labels) {
        checked++;
        const mentioned = new RegExp(`\\bFigure\\s+${n}\\b`).test(
          bodyText.replace(new RegExp(`^Figure\\s+${n}\\.?$`, "gmi"), "")
        );
        if (mentioned) passed++;
        else {
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-FIGURE-002",
            category: "figures",
            severity: "info",
            status: "warning",
            message: `Figure ${n} does not appear to be mentioned in the text.`,
            confidence: 0.7,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-FIGURE-002", checked, passed, false, worst);
    },
  },

  {
    id: "APA-FIGURE-003",
    category: "figures",
    description: "Embedded text inside raster figures requires a visual check.",
    severity: "info",
    applies: (ctx) => ctx.model.imageCount > 0,
    run(ctx) {
      ctx.addIssue({
        ruleId: "APA-FIGURE-003",
        category: "figures",
        severity: "info",
        status: "user_review",
        message: "Check fonts, labels, and readability inside each embedded figure image.",
        explanation: "The formatter can style the Figure number, title, placement, and note, but cannot safely rewrite text baked into a raster image.",
        confidence: 1,
        autoFixable: false,
        userResolutionRequired: true,
      });
      return result("APA-FIGURE-003", ctx.model.imageCount, 0, false, "user_review");
    },
  },
];
