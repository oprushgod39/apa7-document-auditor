import { analyzeDocument } from "../apa/analysis.js";
import type { DocumentModel } from "../docx/model.js";

/**
 * Standalone labels used by the formatter's default student title page.
 * Restrict this fallback to the opening paragraphs so an ordinary body
 * sentence mentioning a professor or university is never removed.
 */
const TITLE_DETAIL_LABEL = /^(?:name|university|institution|subject\s+(?:code\s*(?:and|&)?\s*name|name)|course\s+(?:code|number|name)(?:\s*(?:and|&)?\s*(?:name|number))?|professor|instructor|due\s+date)\s*:?[\s_]*$/i;

function normalized(text: string): string {
  return text.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
}

export function mergeExcludedTitleParagraphIndexes(model: DocumentModel): Set<number> {
  const analysis = analyzeDocument(model);
  const excluded = new Set<number>();

  // Remove the detected student details, but preserve the paper title itself.
  if (analysis.hasTitlePage) {
    const metadataValues = new Set(
      [
        analysis.detectedMetadata.author,
        analysis.detectedMetadata.institution,
        analysis.detectedMetadata.instructor,
        analysis.detectedMetadata.dueDate,
        analysis.detectedMetadata.courseNumber,
        analysis.detectedMetadata.courseName,
        analysis.detectedMetadata.courseNumber && analysis.detectedMetadata.courseName
          ? `${analysis.detectedMetadata.courseNumber} ${analysis.detectedMetadata.courseName}`
          : undefined,
      ]
        .filter((value): value is string => Boolean(value))
        .map(normalized)
    );

    for (const paragraph of model.paragraphs) {
      if (paragraph.index >= analysis.titlePageEnd) break;
      if (paragraph.index === analysis.detectedMetadata.titleParagraphIndex) continue;
      const text = paragraph.text.trim();
      if (!text || TITLE_DETAIL_LABEL.test(text) || metadataValues.has(normalized(text))) {
        excluded.add(paragraph.index);
      }
    }
    return excluded;
  }

  // Malformed source files sometimes contain the default labels without a
  // page break. Remove only those exact labels from the document opening.
  for (const paragraph of model.paragraphs.slice(0, 12)) {
    if (TITLE_DETAIL_LABEL.test(paragraph.text.trim())) excluded.add(paragraph.index);
  }
  return excluded;
}
