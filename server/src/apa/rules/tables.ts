import type { ApaRule } from "../types.js";
import { formatApaTable } from "../../docx/edit.js";
import { result, markDocDirty } from "./util.js";
import { excerptOf } from "../types.js";
import { formatCaptionParagraph } from "./captions.js";

/** N. Table rules — audit and gentle formatting; never restructure tables. */
export const tableRules: ApaRule[] = [
  {
    id: "APA-TABLE-001",
    category: "tables",
    description: `Each table has a bold "Table N" label and an italic title above it.`,
    severity: "warning",
    applies: (ctx) => ctx.model.tables.length > 0,
    run(ctx, fix) {
      const { model } = ctx;
      let checked = 0;
      let passed = 0;
      let fixedCount = 0;
      let worst: "user_review" | null = null;

      for (const table of model.tables) {
        checked++;
        // Look at the two non-empty paragraphs immediately before the table.
        const before = model.paragraphs
          .filter((p) => p.blockIndex < table.blockIndex && !p.isEmpty)
          .slice(-2);
        const numberPara = before.find((p) => /^table\s+\d+\.?$/i.test(p.text.trim()));
        const titlePara = numberPara
          ? before.find((p) => p !== numberPara && p.blockIndex > numberPara.blockIndex)
          : undefined;

        if (!numberPara) {
          worst = "user_review";
          ctx.addIssue({
            ruleId: "APA-TABLE-001",
            category: "tables",
            severity: "warning",
            status: "user_review",
            message: `Table ${table.index + 1} has no "Table N" label above it.`,
            explanation:
              `APA 7 places a bold "Table ${table.index + 1}" label and an italic title on separate lines above each table. The auditor does not invent titles — add them, or confirm the table is presented differently on purpose.`,
            location: { tableIndex: table.index, description: `Table ${table.index + 1} (${table.rows}×${table.cols})` },
            confidence: 0.85,
            autoFixable: false,
            userResolutionRequired: true,
            resolutionOptions: [
              { id: "will_add", label: "I will add the label", description: "I will add the Table N label and title." },
              { id: "intentional", label: "Intentional", description: "This table is intentionally unlabeled (e.g. layout table)." },
            ],
          });
          continue;
        }

        let ok =
          numberPara.runProps.bold === true &&
          numberPara.runProps.italic !== true &&
          numberPara.props.alignment === "left" &&
          titlePara != null &&
          titlePara.runProps.italic === true &&
          titlePara.runProps.bold !== true &&
          titlePara.props.alignment === "left";
        if (fix) {
          formatCaptionParagraph(ctx, numberPara, "label");
          fixedCount++;
          if (titlePara) formatCaptionParagraph(ctx, titlePara, "title");
          ctx.addChange({
            ruleId: "APA-TABLE-001",
            category: "tables",
            location: { paragraphIndex: numberPara.index, excerpt: numberPara.text.trim() },
            before: "source table label/title formatting",
            after: "left-aligned bold table label and left-aligned italic non-bold title",
            reason: "APA 7 table numbers and titles use distinct formatting on separate lines.",
            confidence: 0.95,
          });
          if (!titlePara) {
            worst = "user_review";
            ctx.addIssue({
              ruleId: "APA-TABLE-001",
              category: "tables",
              severity: "warning",
              status: "user_review",
              message: `Table ${table.index + 1} has a label but no separate title/explanation line.`,
              explanation: "Add a descriptive title immediately below the Table N label. The auditor does not invent it.",
              location: { tableIndex: table.index },
              confidence: 0.8,
              autoFixable: false,
              userResolutionRequired: true,
            });
          }
        } else if (!ok) {
          ctx.addIssue({
            ruleId: "APA-TABLE-001",
            category: "tables",
            severity: "warning",
            status: "fail",
            message: `Table ${table.index + 1} label/title formatting does not follow APA (left-aligned bold number; left-aligned italic non-bold title).`,
            location: { tableIndex: table.index },
            confidence: 0.9,
            autoFixable: titlePara != null,
            userResolutionRequired: titlePara == null,
          });
        }
        if (ok) passed++;
      }
      return result("APA-TABLE-001", checked, passed, fixedCount > 0, worst);
    },
  },

  {
    id: "APA-TABLE-003",
    category: "tables",
    description: "Tables use APA horizontal rules, repeating headers, double spacing, and unsplit rows.",
    severity: "warning",
    applies: (ctx) => ctx.model.tables.length > 0,
    run(ctx, fix) {
      if (fix) {
        for (const table of ctx.model.tables) {
          formatApaTable(ctx.model.documentXml, table.el, {
            font: ctx.req.font,
            halfPoints: ctx.req.fontSizePt * 2,
          });
        }
        markDocDirty(ctx);
        ctx.addChange({
          ruleId: "APA-TABLE-003",
          category: "tables",
          location: { description: "All document tables" },
          before: "source borders, row pagination, and cell spacing",
          after: "APA horizontal rules, repeating header rows, unsplit rows, and double-spaced cells",
          reason: "APA tables avoid vertical rules and should remain readable across page boundaries.",
          confidence: 0.95,
          documentWide: true,
        });
        return result("APA-TABLE-003", ctx.model.tables.length, 0, true, null);
      }
      return result("APA-TABLE-003", ctx.model.tables.length, ctx.model.tables.length, false, null);
    },
  },

  {
    id: "APA-TABLE-002",
    category: "tables",
    description: "Each table is referred to in the text (e.g. “Table 1”).",
    severity: "info",
    applies: (ctx) => ctx.model.tables.length > 0,
    run(ctx) {
      const { model, analysis } = ctx;
      let checked = 0;
      let passed = 0;
      let worst: "warning" | null = null;
      const bodyText = model.paragraphs
        .filter(
          (p) =>
            !p.insideTable &&
            p.index >= analysis.bodyStartIndex &&
            (analysis.referencesHeadingIndex == null ||
              p.index < analysis.referencesHeadingIndex)
        )
        .map((p) => p.text)
        .join("\n");
      for (const table of model.tables) {
        checked++;
        const n = table.index + 1;
        const mentionRe = new RegExp(`\\bTable\\s+${n}\\b`);
        // Ignore the caption line itself by requiring the mention within prose
        const mentioned = mentionRe.test(
          bodyText.replace(new RegExp(`^Table\\s+${n}\\.?$`, "gmi"), "")
        );
        if (mentioned) passed++;
        else {
          worst = "warning";
          ctx.addIssue({
            ruleId: "APA-TABLE-002",
            category: "tables",
            severity: "info",
            status: "warning",
            message: `Table ${n} does not appear to be mentioned in the text.`,
            explanation: "APA 7 requires each table to be called out in the narrative (e.g. “as shown in Table 1”).",
            location: { tableIndex: table.index },
            confidence: 0.7,
            autoFixable: false,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-TABLE-002", checked, passed, false, worst);
    },
  },
];
