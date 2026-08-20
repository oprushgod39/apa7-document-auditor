import type { RuleContext } from "../types.js";
import type { ParagraphModel } from "../../docx/model.js";
import {
  setParagraphAlignment,
  setParagraphContextualSpacing,
  setParagraphIndent,
  setParagraphRunFonts,
  setParagraphSpacing,
  setParagraphKeepNext,
  setRunBold,
  setRunColorBlack,
  setRunItalic,
  setRunUnderlineNone,
} from "../../docx/edit.js";
import { childrenW } from "../../docx/xml.js";
import { markDocDirty } from "./util.js";

/** Apply explicit APA formatting at both paragraph and run levels. */
export function formatCaptionParagraph(
  ctx: RuleContext,
  p: ParagraphModel,
  kind: "label" | "title"
): void {
  const doc = ctx.model.documentXml;
  setParagraphAlignment(doc, p.el, "left");
  setParagraphIndent(doc, p.el, { firstLine: null, hanging: null, left: 0 });
  setParagraphSpacing(doc, p.el, { before: 0, after: 0, line: 480, lineRule: "auto" });
  setParagraphContextualSpacing(doc, p.el);
  setParagraphKeepNext(doc, p.el, true);
  setParagraphRunFonts(doc, p.el, ctx.req.font, ctx.req.fontSizePt * 2);
  for (const r of childrenW(p.el, "r")) {
    setRunBold(doc, r, kind === "label");
    setRunItalic(doc, r, kind === "title");
    setRunColorBlack(doc, r);
    setRunUnderlineNone(doc, r);
  }
  markDocDirty(ctx);
}
