import {
  DEFAULT_APPROVED_FONTS,
  TWIPS_PER_INCH,
  APA_DOUBLE_LINE,
  APA_FIRST_LINE_INDENT,
  APA_HANGING_INDENT,
  type DocumentSettings,
  type EffectiveRequirements,
  type InstructorRequirements,
  type PaperType,
} from "./types.js";

/**
 * Parse pasted instructor/assignment requirements into structured overrides.
 * Only unambiguous statements are interpreted; everything else is preserved
 * verbatim for the user to see under "uninterpreted".
 */
export function parseInstructorRequirements(rawText: string): InstructorRequirements {
  const req: InstructorRequirements = { rawText, uninterpreted: [] };
  const lines = rawText
    .split(/\r?\n|[;•]/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  for (const line of lines) {
    let interpreted = false;
    const lower = line.toLowerCase();

    const fontMatch =
      /(times new roman|calibri|arial|georgia|lucida sans unicode|computer modern)\s*,?\s*(\d{1,2})?\s*(?:pt|point)?/i.exec(
        line
      );
    if (fontMatch && /font|times new roman|calibri|arial|georgia/i.test(lower)) {
      req.font = titleCaseFont(fontMatch[1]!);
      if (fontMatch[2]) req.fontSizePt = Number.parseInt(fontMatch[2], 10);
      interpreted = true;
    }
    const sizeOnly = /(\d{1,2})\s*(?:pt|point)\b/i.exec(line);
    if (!fontMatch && sizeOnly && /font|size/i.test(lower)) {
      req.fontSizePt = Number.parseInt(sizeOnly[1]!, 10);
      interpreted = true;
    }

    if (/\babstract\b/.test(lower)) {
      if (/\bno abstract\b|without an? abstract|abstract (is )?not (required|needed)/.test(lower)) {
        req.abstractRequired = false;
        interpreted = true;
      } else if (/required|include|must|needs?\b/.test(lower)) {
        req.abstractRequired = true;
        interpreted = true;
      }
    }

    if (/running head/.test(lower)) {
      if (/\bno running head\b|not (required|needed)/.test(lower)) {
        req.runningHeadRequired = false;
      } else {
        req.runningHeadRequired = true;
      }
      interpreted = true;
    }

    const minRefs =
      /(?:minimum|at least|no fewer than)\s+(?:of\s+)?(\d+|five|four|three|six|seven|eight|nine|ten)\s+(?:references|sources|citations)/i.exec(
        line
      ) ||
      /(\d+)\+?\s+(?:references|sources)\s+(?:minimum|required)/i.exec(line);
    if (minRefs) {
      req.minReferences = wordToNumber(minRefs[1]!);
      interpreted = true;
    }

    if (!interpreted) req.uninterpreted.push(line);
  }
  return req;
}

function titleCaseFont(s: string): string {
  return s
    .split(/\s+/)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ")
    .replace(/\bnew\b/i, "New")
    .replace(/\bRoman\b/i, "Roman");
}

function wordToNumber(s: string): number {
  const words: Record<string, number> = {
    three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  };
  const n = Number.parseInt(s, 10);
  return Number.isFinite(n) ? n : (words[s.toLowerCase()] ?? 5);
}

/** Layer instructor requirements over the APA 7 baseline. */
export function effectiveRequirements(
  settings: DocumentSettings
): EffectiveRequirements {
  const inst = settings.instructor;
  const overrides: string[] = [];

  let font = "Times New Roman";
  let fontSizePt = 12;
  if (inst.font) {
    font = inst.font;
    const approved = DEFAULT_APPROVED_FONTS.find(
      (f) => f.name.toLowerCase() === inst.font!.toLowerCase()
    );
    fontSizePt = inst.fontSizePt ?? approved?.sizePt ?? 12;
    overrides.push(`Instructor override: font ${font} ${fontSizePt} pt`);
  } else if (inst.fontSizePt) {
    fontSizePt = inst.fontSizePt;
    overrides.push(`Instructor override: font size ${fontSizePt} pt`);
  }

  const runningHeadDefault: boolean = settings.paperType === "professional";
  let runningHeadRequired = runningHeadDefault;
  if (inst.runningHeadRequired != null) {
    runningHeadRequired = inst.runningHeadRequired;
    if (inst.runningHeadRequired !== runningHeadDefault) {
      overrides.push(
        `Instructor override: running head ${inst.runningHeadRequired ? "required" : "not required"}`
      );
    }
  }

  let abstractRequired: boolean | null =
    settings.paperType === "professional" ? null : null;
  if (inst.abstractRequired != null) {
    abstractRequired = inst.abstractRequired;
    overrides.push(
      `Instructor override: abstract ${inst.abstractRequired ? "required" : "not required"}`
    );
  }

  if (inst.minReferences != null) {
    overrides.push(`Instructor requirement: minimum ${inst.minReferences} references`);
  }

  return {
    font,
    fontSizePt,
    approvedFonts: DEFAULT_APPROVED_FONTS,
    marginTwips: TWIPS_PER_INCH,
    lineSpacing: APA_DOUBLE_LINE,
    firstLineIndentTwips: APA_FIRST_LINE_INDENT,
    hangingIndentTwips: APA_HANGING_INDENT,
    abstractRequired,
    runningHeadRequired,
    minReferences: inst.minReferences ?? null,
    overrides,
  };
}

export function defaultSettings(paperType: PaperType = "student"): DocumentSettings {
  return {
    paperType,
    mode: "format_verify",
    preserveWording: true,
    fixCitationMechanics: true,
    verifyMetadata: true,
    metadata: {},
    instructor: { uninterpreted: [] },
  };
}
