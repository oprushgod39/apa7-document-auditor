import type { ApaRule } from "../types.js";
import { result, markDocDirty } from "./util.js";
import { setRunBold, setRunItalic } from "../../docx/edit.js";
import { childrenW } from "../../docx/xml.js";
import { excerptOf } from "../types.js";

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

        let ok = true;
        if (numberPara.runProps.bold !== true) {
          if (fix) {
            for (const r of childrenW(numberPara.el, "r")) {
              setRunBold(model.documentXml, r, true);
            }
            markDocDirty(ctx);
            fixedCount++;
            ctx.addChange({
              ruleId: "APA-TABLE-001",
              category: "tables",
              location: { paragraphIndex: numberPara.index, excerpt: numberPara.text.trim() },
              before: `"${numberPara.text.trim()}" not bold`,
              after: `"${numberPara.text.trim()}" bold`,
              reason: `APA 7 table numbers ("Table N") are bold.`,
              confidence: 0.95,
            });
          } else {
            ok = false;
          }
        }
        if (titlePara && titlePara.runProps.italic !== true) {
          if (fix) {
            for (const r of childrenW(titlePara.el, "r")) {
              setRunItalic(model.documentXml, r, true);
            }
            markDocDirty(ctx);
            fixedCount++;
            ctx.addChange({
              ruleId: "APA-TABLE-001",
              category: "tables",
              location: { paragraphIndex: titlePara.index, excerpt: excerptOf(titlePara.text) },
              before: "table title not italic",
              after: "table title italic",
              reason: "APA 7 table titles are italic, in title case.",
              confidence: 0.9,
            });
          } else {
            ok = false;
          }
        }
        if (ok && fixedCount === 0) passed++;
        if (!fix && !ok) {
          ctx.addIssue({
            ruleId: "APA-TABLE-001",
            category: "tables",
            severity: "warning",
            status: "fail",
            message: `Table ${table.index + 1} label/title formatting does not follow APA (bold number, italic title).`,
            location: { tableIndex: table.index },
            confidence: 0.9,
            autoFixable: true,
            userResolutionRequired: false,
          });
        }
      }
      return result("APA-TABLE-001", checked, passed, fixedCount > 0, worst);
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
