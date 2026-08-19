import type { DocumentModel, ParagraphModel } from "../docx/model.js";
import { parseCitations, type ParsedCitation } from "./citations/parser.js";
import { parseReference, type ParsedReference } from "./references/parser.js";
import {
  matchCitationsToReferences,
  type MatchingResult,
} from "./references/matcher.js";
import {
  classifyHeadings,
  type ClassifiedHeading,
} from "./headings/classifier.js";
import type { TitlePageMetadata } from "./types.js";

/**
 * One-pass structural analysis shared by all rules. Built once per document
 * model; cached per processing session so the document is not re-parsed
 * repeatedly.
 */
export interface DocumentAnalysis {
  /** Paragraph indexes belonging to the title page (before first page break). */
  titlePageEnd: number; // exclusive; 0 when no title page detected
  hasTitlePage: boolean;
  detectedMetadata: TitlePageMetadata & { titleParagraphIndex?: number };
  /** Index of the "Abstract" heading paragraph, if present. */
  abstractHeadingIndex: number | null;
  abstractBodyIndexes: number[];
  keywordsParagraphIndex: number | null;
  /** Index of the References heading paragraph, if present. */
  referencesHeadingIndex: number | null;
  referenceEntryIndexes: number[];
  /** First body paragraph (after title page and abstract). */
  bodyStartIndex: number;
  headings: ClassifiedHeading[];
  citations: ParsedCitation[];
  references: ParsedReference[];
  matching: MatchingResult;
  /** Candidate block quotations: 40+ word quoted passages. */
  longQuoteCandidates: { paragraphIndex: number; words: number; isIndented: boolean }[];
}

const REFERENCES_HEADINGS = /^(references|reference list|works cited|bibliography|reference)$/i;

function findFirstPageBreak(model: DocumentModel): number | null {
  for (const p of model.paragraphs) {
    if (p.hasPageBreakAfterInRuns || (p.props.pageBreakBefore && p.index > 0)) {
      return p.props.pageBreakBefore ? p.index : p.index + 1;
    }
  }
  return null;
}

const DATE_RE =
  /^(?:(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December),?\s+\d{4}|\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})$/i;
const COURSE_RE = /^[A-Z]{2,6}[- ]?\d{2,4}[A-Z]?\b/;
const INSTITUTION_RE =
  /(university|college|institute|school|academy|department of)/i;
const INSTRUCTOR_RE = /^(dr\.|prof\.?|professor|instructor|mr\.|ms\.|mrs\.)\s/i;

function detectTitlePageMetadata(
  model: DocumentModel,
  end: number
): DocumentAnalysis["detectedMetadata"] {
  const meta: DocumentAnalysis["detectedMetadata"] = {};
  const paras = model.paragraphs.slice(0, end).filter((p) => !p.isEmpty);
  for (const p of paras) {
    const text = p.text.trim();
    if (!meta.title && text.length >= 4 && text.length <= 200) {
      // First substantial non-empty line is the best title candidate.
      meta.title = text;
      meta.titleParagraphIndex = p.index;
      continue;
    }
    if (!meta.dueDate && DATE_RE.test(text)) {
      meta.dueDate = text;
      continue;
    }
    if (!meta.courseNumber && COURSE_RE.test(text)) {
      const m = COURSE_RE.exec(text)!;
      meta.courseNumber = m[0].trim();
      const rest = text.slice(m[0].length).replace(/^[:\-–\s]+/, "").trim();
      if (rest) meta.courseName = rest;
      continue;
    }
    if (!meta.institution && INSTITUTION_RE.test(text) && text.length < 120) {
      meta.institution = text;
      continue;
    }
    if (!meta.instructor && INSTRUCTOR_RE.test(text)) {
      meta.instructor = text;
      continue;
    }
    if (
      !meta.author &&
      meta.title &&
      /^[A-ZÀ-Þ][\w'’\-À-ÿ]+(?:\s[A-ZÀ-Þ][\w.'’\-À-ÿ]+){0,4}$/.test(text) &&
      text.length < 80
    ) {
      meta.author = text;
    }
  }
  return meta;
}

export function analyzeDocument(model: DocumentModel): DocumentAnalysis {
  const paras = model.paragraphs;

  // --- Title page ------------------------------------------------------
  const firstBreak = findFirstPageBreak(model);
  let titlePageEnd = 0;
  let hasTitlePage = false;
  if (firstBreak != null && firstBreak <= 30) {
    const chunk = paras.slice(0, firstBreak).filter((p) => !p.isEmpty);
    // A title page is a handful of short lines with no long prose.
    const longProse = chunk.filter((p) => p.text.trim().split(/\s+/).length > 30);
    if (chunk.length >= 1 && chunk.length <= 12 && longProse.length === 0) {
      hasTitlePage = true;
      titlePageEnd = firstBreak;
    }
  }
  const detectedMetadata = hasTitlePage
    ? detectTitlePageMetadata(model, titlePageEnd)
    : {};

  // --- Abstract --------------------------------------------------------
  let abstractHeadingIndex: number | null = null;
  const abstractBodyIndexes: number[] = [];
  let keywordsParagraphIndex: number | null = null;
  for (let i = titlePageEnd; i < Math.min(paras.length, titlePageEnd + 8); i++) {
    const p = paras[i]!;
    if (p.isEmpty) continue;
    if (/^abstract$/i.test(p.text.trim())) {
      abstractHeadingIndex = i;
      for (let j = i + 1; j < paras.length; j++) {
        const q = paras[j]!;
        if (q.isEmpty) continue;
        if (/^keywords?\s*:/i.test(q.text.trim())) {
          keywordsParagraphIndex = j;
          break;
        }
        if (q.hasPageBreakAfterInRuns || q.props.pageBreakBefore) break;
        if (abstractBodyIndexes.length >= 3) break;
        abstractBodyIndexes.push(j);
      }
    }
    break; // only inspect the first non-empty paragraph after the title page
  }

  // --- References section ---------------------------------------------
  let referencesHeadingIndex: number | null = null;
  for (let i = paras.length - 1; i >= 0; i--) {
    const p = paras[i]!;
    if (p.insideTable) continue;
    if (REFERENCES_HEADINGS.test(p.text.trim())) {
      // Require at least one following non-empty paragraph that looks like a
      // reference (contains a (Year) or a URL) to avoid false positives.
      const following = paras.slice(i + 1).filter((q) => !q.isEmpty);
      const looksLikeRefs = following.some((q) =>
        /\((?:1[6-9]|20)\d{2}[a-z]?\)|\(n\.d\.\)|https?:\/\//.test(q.text)
      );
      if (looksLikeRefs || following.length === 0) {
        referencesHeadingIndex = i;
        break;
      }
    }
  }
  const referenceEntryIndexes: number[] = [];
  if (referencesHeadingIndex != null) {
    for (let i = referencesHeadingIndex + 1; i < paras.length; i++) {
      const p = paras[i]!;
      if (p.isEmpty) continue;
      // Stop at a subsequent heading-like paragraph (e.g. Appendix)
      if (/^(appendix|appendices|footnotes|tables|figures)\b/i.test(p.text.trim()) && p.text.trim().length < 30) break;
      referenceEntryIndexes.push(i);
    }
  }

  // --- Body start ------------------------------------------------------
  let bodyStartIndex = titlePageEnd;
  if (keywordsParagraphIndex != null) bodyStartIndex = keywordsParagraphIndex + 1;
  else if (abstractBodyIndexes.length > 0) {
    bodyStartIndex = abstractBodyIndexes[abstractBodyIndexes.length - 1]! + 1;
  } else if (abstractHeadingIndex != null) bodyStartIndex = abstractHeadingIndex + 1;

  // --- Headings --------------------------------------------------------
  const headings = classifyHeadings(model, {
    bodyStartIndex,
    referencesHeadingIndex,
  });

  // --- Citations (body only, excluding reference list) -----------------
  const citations: ParsedCitation[] = [];
  const refSet = new Set(referenceEntryIndexes);
  for (const p of paras) {
    if (p.isEmpty || refSet.has(p.index)) continue;
    if (p.index < titlePageEnd) continue;
    if (referencesHeadingIndex != null && p.index >= referencesHeadingIndex) continue;
    citations.push(...parseCitations(p.text, p.index));
  }

  // --- References ------------------------------------------------------
  const references: ParsedReference[] = referenceEntryIndexes.map((i) =>
    parseReference(paras[i]!.text, i)
  );

  const matching = matchCitationsToReferences(citations, references);

  // --- Long quotes -----------------------------------------------------
  const longQuoteCandidates: DocumentAnalysis["longQuoteCandidates"] = [];
  for (const p of paras) {
    if (p.isEmpty || refSet.has(p.index) || p.index < bodyStartIndex) continue;
    const quoteMatch = /[“"]([^”"]{150,})[”"]/.exec(p.text);
    if (quoteMatch) {
      const words = quoteMatch[1]!.split(/\s+/).length;
      if (words >= 40) {
        longQuoteCandidates.push({
          paragraphIndex: p.index,
          words,
          isIndented: (p.props.leftIndent ?? 0) >= 700,
        });
      }
    } else if (
      (p.props.leftIndent ?? 0) >= 700 &&
      p.text.trim().split(/\s+/).length >= 40 &&
      !p.hasNumbering
    ) {
      // Already-indented block quote without quotation marks.
      longQuoteCandidates.push({
        paragraphIndex: p.index,
        words: p.text.trim().split(/\s+/).length,
        isIndented: true,
      });
    }
  }

  return {
    titlePageEnd,
    hasTitlePage,
    detectedMetadata,
    abstractHeadingIndex,
    abstractBodyIndexes,
    keywordsParagraphIndex,
    referencesHeadingIndex,
    referenceEntryIndexes,
    bodyStartIndex,
    headings,
    citations,
    references,
    matching,
    longQuoteCandidates,
  };
}

/** Convenience: paragraph by index or throw. */
export function paraAt(model: DocumentModel, index: number): ParagraphModel {
  const p = model.paragraphs[index];
  if (!p) throw new Error(`Paragraph ${index} out of range`);
  return p;
}
