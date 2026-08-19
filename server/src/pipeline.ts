import { DocxPackage, verifyDocxIntegrity } from "./docx/package.js";
import {
  buildDocumentModel,
  contentFingerprint,
  type ContentFingerprint,
} from "./docx/model.js";
import {
  ensureParagraphStyle,
  setParagraphStyle,
  removeParagraphStyle,
} from "./docx/edit.js";
import { runEngine, allRules } from "./apa/engine.js";
import { analyzeDocument } from "./apa/analysis.js";
import { HEADING_SPECS } from "./apa/rules/headings.js";
import type { Change, RuleCategory } from "./apa/types.js";
import { buildReport, issueKey, type ComplianceReport } from "./audit/auditor.js";
import { getProvider } from "./verify/crossref.js";
import { CrossrefProvider } from "./verify/crossref.js";
import type { VerificationResult } from "./verify/provider.js";
import {
  readOriginal,
  writeOutput,
  STAGE_LABELS,
  type Session,
  type ProcessingStageKey,
} from "./store/sessions.js";
import { AppError, Errors } from "./errors.js";
import { log } from "./logging.js";

const ruleCategoryById = new Map<string, RuleCategory>(
  allRules().map((r) => [r.id, r.category])
);

function stage(session: Session, key: ProcessingStageKey, status: "running" | "done" | "skipped" | "failed"): void {
  const existing = session.stages.find((s) => s.key === key);
  if (existing) existing.status = status;
  else session.stages.push({ key, label: STAGE_LABELS[key], status });
}

/** Collapse text to an alphanumeric stream for content comparison. */
function collapse(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Content-preservation guard. Verifies the formatter did not change
 * substantive text: structural counts must match exactly, and any text
 * difference must be explained by the recorded change log (citation
 * mechanics, DOI presentation, heading label, inserted metadata).
 */
export function assertContentPreserved(
  before: ContentFingerprint,
  after: ContentFingerprint,
  changes: Change[]
): void {
  const problems: string[] = [];
  if (before.imageCount !== after.imageCount) {
    problems.push(`images: ${before.imageCount} → ${after.imageCount}`);
  }
  if (before.tableShapes.join(",") !== after.tableShapes.join(",")) {
    problems.push("table structure changed");
  }
  if (before.hyperlinkCount !== after.hyperlinkCount) {
    problems.push(`hyperlinks: ${before.hyperlinkCount} → ${after.hyperlinkCount}`);
  }
  if (before.footnoteCount !== after.footnoteCount) {
    problems.push(`footnotes: ${before.footnoteCount} → ${after.footnoteCount}`);
  }
  if (before.endnoteCount !== after.endnoteCount) {
    problems.push(`endnotes: ${before.endnoteCount} → ${after.endnoteCount}`);
  }
  if (before.equationCount !== after.equationCount) {
    problems.push(`equations: ${before.equationCount} → ${after.equationCount}`);
  }

  // Paragraph-level multiset comparison. Order changes (e.g. alphabetizing
  // the reference list) are allowed; content loss or unexplained edits are
  // not. Textual changes recorded in the change log are replayed onto the
  // "before" paragraphs so that legitimate mechanics fixes reconcile.
  let expectedParas = before.paragraphs.map(collapse).filter((t) => t.length > 0);
  const actualParas = after.paragraphs.map(collapse).filter((t) => t.length > 0);

  for (const c of changes) {
    if (c.excluded) continue;
    // Replay either the literal before/after, or — when the change log holds
    // a description — the first quoted values inside it (e.g.
    // `"Bibliography", not bold` → `"References", bold`).
    const candidates: [string, string][] = [[collapse(c.before), collapse(c.after)]];
    const qb = /[“"]([^”"]+)[”"]/.exec(c.before);
    const qa = /[“"]([^”"]+)[”"]/.exec(c.after);
    if (qb && qa) candidates.push([collapse(qb[1]!), collapse(qa[1]!)]);
    for (const [b, a] of candidates) {
      if (b.length === 0 || a === b) continue;
      const i = expectedParas.findIndex((t) => t.includes(b));
      if (i >= 0) {
        expectedParas[i] = expectedParas[i]!.replace(b, a);
        break;
      }
    }
  }

  const titlePageInsertion = changes.some(
    (c) =>
      !c.excluded &&
      (c.ruleId === "APA-TITLE-001" || c.ruleId === "APA-TITLE-003")
  );

  const remaining = [...actualParas];
  const missing: string[] = [];
  for (const exp of expectedParas) {
    const i = remaining.indexOf(exp);
    if (i >= 0) remaining.splice(i, 1);
    else missing.push(exp);
  }
  if (missing.length > 0) {
    problems.push(
      `document text changed beyond recorded modifications (${missing.length} paragraph(s) altered or lost)`
    );
  }
  // Extra paragraphs in the output are only legitimate when title-page
  // metadata was inserted from user-provided values.
  if (remaining.length > 0 && !titlePageInsertion) {
    problems.push(
      `unexpected content appeared in the output (${remaining.length} paragraph(s))`
    );
  }

  if (problems.length > 0) {
    log.error("content-preservation guard tripped", { problems });
    throw new AppError(
      "PROCESSING_FAILED",
      "Formatting was aborted because a safety check detected an unexpected content change. Your original document is unaffected.",
      500
    );
  }
}

/** Full processing pipeline for a session. Never mutates the original file. */
export async function processSession(session: Session): Promise<void> {
  const started = Date.now();
  session.status = "processing";
  session.stages = [];
  session.errorMessage = null;
  try {
    const mode = session.settings.mode;

    // 1. Read + parse the original (a fresh in-memory copy every run).
    stage(session, "read", "running");
    const originalBuffer = await readOriginal(session);
    const pkg = await DocxPackage.load(originalBuffer);
    const model = await buildDocumentModel(pkg);
    stage(session, "read", "done");

    stage(session, "structure", "running");
    const beforeFp = contentFingerprint(model);
    const analysis = analyzeDocument(model);
    session.cachedAnalysis = analysis;
    stage(session, "structure", "done");
    stage(session, "headings", "done");
    stage(session, "page_format", "done");
    stage(session, "citations", "done");
    stage(session, "references", "done");

    const fix = mode !== "check";

    // 2. Format (or check-only) run.
    stage(session, "apply", fix ? "running" : "skipped");
    const formatRun = await runEngine(model, session.settings, {
      fix,
      analysis,
      stage: "format",
      disabledRules: session.disabledRules,
    });
    let changes: Change[] = [...formatRun.changes];

    // 3. Apply user-forced heading levels (from prior resolutions).
    if (fix && session.forcedHeadings.size > 0) {
      for (const [paragraphIndex, level] of session.forcedHeadings) {
        const p = model.paragraphs[paragraphIndex];
        if (!p) continue;
        if (level === 0) {
          removeParagraphStyle(p.el);
        } else {
          const spec = HEADING_SPECS[level];
          if (!spec) continue;
          if (model.stylesXml) {
            ensureParagraphStyle(model.stylesXml, spec.styleId, spec.name);
            model.pkg.markDirty("word/styles.xml");
          }
          setParagraphStyle(model.documentXml, p.el, spec.styleId);
        }
        model.pkg.markDirty("word/document.xml");
        changes.push({
          id: `resolution-${paragraphIndex}`,
          ruleId: "APA-HEAD-002",
          category: "headings",
          location: { paragraphIndex },
          before: "Heading level unresolved",
          after: level === 0 ? "Confirmed as normal paragraph" : `Confirmed as Level ${level} heading`,
          reason: "Resolved by you.",
          confidence: 1,
          stage: "resolution",
          timestamp: new Date().toISOString(),
        });
      }
    }
    if (fix) stage(session, "apply", "done");

    // 4. Save modified package and verify integrity.
    stage(session, "prepare_output", "running");
    let outputBuffer: Buffer;
    if (fix) {
      outputBuffer = await pkg.save();
      await verifyDocxIntegrity(outputBuffer);
    } else {
      outputBuffer = originalBuffer;
    }

    // 5. Content-preservation guard on the regenerated document.
    const outPkg = await DocxPackage.load(outputBuffer);
    const outModel = await buildDocumentModel(outPkg);
    if (fix) {
      assertContentPreserved(beforeFp, contentFingerprint(outModel), changes);
    }
    stage(session, "prepare_output", "done");

    // 6. External metadata verification (graceful degradation).
    let verification: VerificationResult[] | null = null;
    const outAnalysis = analyzeDocument(outModel);
    if (
      mode === "format_verify" &&
      session.settings.verifyMetadata &&
      outAnalysis.references.length > 0
    ) {
      stage(session, "verify_metadata", "running");
      const provider = getProvider();
      if (provider instanceof CrossrefProvider) provider.resetBudget();
      verification = [];
      for (const ref of outAnalysis.references) {
        verification.push(await provider.verify(ref));
      }
      stage(session, "verify_metadata", "done");
    } else {
      stage(session, "verify_metadata", "skipped");
    }

    // 7. Independent audit of the OUTPUT document (check-only, no fixes).
    stage(session, "audit", "running");
    const auditRun = await runEngine(outModel, session.settings, {
      fix: false,
      auditOnly: true,
      analysis: outAnalysis,
      disabledRules: session.disabledRules,
    });

    // Fold verification results into audit issues.
    if (verification) {
      for (const v of verification) {
        if (v.status === "verified" || v.status === "unverified") continue;
        const ref = outAnalysis.references.find(
          (r) => r.paragraphIndex === v.referenceIndex
        );
        if (!ref) continue;
        const isMismatch = v.status === "mismatch";
        auditRun.issues.push({
          id: `verify-${v.referenceIndex}`,
          ruleId: "APA-REFERENCE-VERIFY",
          category: "references",
          severity: isMismatch ? "warning" : "info",
          status: isMismatch ? "user_review" : v.status === "provider_unavailable" ? "unverified" : "warning",
          message: isMismatch
            ? `Reference metadata differs from the ${v.provider} record.`
            : v.status === "provider_unavailable"
              ? "External metadata verification temporarily unavailable."
              : `Reference verified with minor differences (${Math.round(v.confidence * 100)}%).`,
          explanation: v.differences
            ?.map(
              (d) => `${d.field} — document: “${d.documentValue}” / verified: “${d.verifiedValue}”`
            )
            .join("; ") ?? v.note,
          location: { paragraphIndex: v.referenceIndex, excerpt: ref.raw.slice(0, 80) },
          confidence: v.confidence,
          autoFixable: false,
          userResolutionRequired: isMismatch,
          resolutionOptions: isMismatch
            ? [
                { id: "document_correct", label: "My entry is correct", description: "Keep the document values." },
                { id: "will_fix", label: "I will correct it", description: "I will update the reference to the verified values." },
              ]
            : undefined,
        });
      }
    }

    const report = buildReport(
      auditRun,
      changes,
      verification,
      session.resolutions,
      new Map([...ruleCategoryById, ["APA-REFERENCE-VERIFY", "references" as RuleCategory]])
    );
    stage(session, "audit", "done");

    if (fix) await writeOutput(session, outputBuffer);
    session.report = report;
    session.changes = changes;
    session.verification = verification;
    session.status = "ready";
    log.info("processing complete", {
      sessionId: session.id,
      mode,
      ms: Date.now() - started,
      rules: auditRun.ruleResults.length,
      issues: auditRun.issues.length,
      changes: changes.length,
    });
  } catch (err) {
    session.status = "error";
    session.errorMessage =
      err instanceof AppError ? err.userMessage : Errors.internal().userMessage;
    for (const s of session.stages) {
      if (s.status === "running") s.status = "failed";
    }
    log.error("processing failed", {
      sessionId: session.id,
      code: err instanceof AppError ? err.code : "INTERNAL",
      error: err instanceof Error ? err.message : String(err),
    });
    if (!(err instanceof AppError)) throw Errors.internal();
    throw err;
  }
}

/** Record a user resolution; heading levels are captured for regeneration. */
export function applyResolution(
  session: Session,
  key: string,
  optionId: string,
  note?: string
): void {
  session.resolutions.set(key, { optionId, note });
  const report = session.report;
  if (!report) return;
  const issue = report.issues.find((i) => issueKey(i) === key);
  if (issue) {
    issue.resolution = { optionId, note };
    (issue as { resolved: boolean }).resolved = true;
    if (
      issue.ruleId === "APA-HEAD-002" ||
      issue.ruleId === "APA-HEAD-003"
    ) {
      const idx = issue.location?.paragraphIndex;
      const m = /^level([1-5])$/.exec(optionId);
      if (idx != null && m) {
        session.forcedHeadings.set(idx, Number.parseInt(m[1]!, 10));
      } else if (idx != null && optionId === "normal") {
        session.forcedHeadings.set(idx, 0);
      }
    }
  }
  // Recompute report state and the supplementary percentage.
  const unresolved = report.issues.filter(
    (i) => i.userResolutionRequired && i.resolution == null
  );
  report.unresolvedCount = unresolved.length;
  report.state = unresolved.length === 0 ? "apa_validated" : "review_required";
  if (report.applicableRules > 0) {
    const openRules = new Set(
      report.issues
        .filter(
          (i) =>
            (i.userResolutionRequired && i.resolution == null) ||
            (i.severity === "error" && i.status === "fail")
        )
        .map((i) => i.ruleId)
    );
    report.rulesResolvedPercent = Math.round(
      ((report.applicableRules - openRules.size) / report.applicableRules) * 100
    );
  }
}
