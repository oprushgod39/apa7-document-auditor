import type { ApaRule } from "../types.js";
import { result, markDocDirty } from "./util.js";
import { setRunBold, setRunItalic } from "../../docx/edit.js";
import { childrenW } from "../../docx/xml.js";

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
        if (numberPara.runProps.bold !== true && fix) {
          for (const r of childrenW(numberPara.el, "r")) {
            setRunBold(model.documentXml, r, true);
          }
          markDocDirty(ctx);
          fixedCount++;
          ctx.addChange({
            ruleId: "APA-FIGURE-001",
            category: "figures",
            location: { paragraphIndex: numberPara.index, excerpt: numberPara.text.trim() },
            before: `"${numberPara.text.trim()}" not bold`,
            after: `"${numberPara.text.trim()}" bold`,
            reason: `APA 7 figure numbers ("Figure N") are bold.`,
            confidence: 0.95,
          });
        } else if (numberPara.runProps.bold === true) {
          passed++;
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
];
