import { Router, type Request, type Response, type NextFunction } from "express";
import multer from "multer";
import { z } from "zod";
import path from "node:path";
import { config } from "../config.js";
import { AppError, Errors } from "../errors.js";
import { DocxPackage } from "../docx/package.js";
import { buildDocumentModel } from "../docx/model.js";
import { analyzeDocument } from "../apa/analysis.js";
import { parseInstructorRequirements, defaultSettings } from "../apa/requirements.js";
import {
  createSession,
  getSession,
  saveSession,
  deleteSession,
  readOriginal,
  readOutput,
  type Session,
} from "../store/sessions.js";
import { processSession, applyResolution } from "../pipeline.js";
import { runInBackground } from "../run_background.js";
import { renderReportHtml } from "../audit/report_html.js";
import { log } from "../logging.js";
import { mergeDocuments } from "../merge/merge.js";
import { appendixSourceWordCount, countDocumentWords } from "../merge/word_count.js";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: config.maxUploadBytes, files: 1 },
});

const mergeUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: config.maxUploadBytes, files: 30 },
});

const SettingsSchema = z.object({
  paperType: z.enum(["student", "professional"]).default("student"),
  mode: z.enum(["check", "format", "format_verify"]).default("format_verify"),
  preserveWording: z.boolean().default(true),
  fixCitationMechanics: z.boolean().default(true),
  verifyMetadata: z.boolean().default(true),
  metadata: z
    .object({
      title: z.string().max(300).optional(),
      author: z.string().max(200).optional(),
      institution: z.string().max(200).optional(),
      courseNumber: z.string().max(100).optional(),
      courseName: z.string().max(200).optional(),
      instructor: z.string().max(200).optional(),
      dueDate: z.string().max(100).optional(),
      authorNote: z.string().max(2000).optional(),
      runningHead: z.string().max(60).optional(),
    })
    .default({}),
  instructorRequirements: z.string().max(8000).default(""),
});

const ResolveSchema = z.object({
  issueKey: z.string().min(1).max(1000),
  optionId: z.string().min(1).max(100),
  note: z.string().max(1000).optional(),
});

const RulesSchema = z.object({
  disabledRules: z.array(z.string().max(60)).max(200),
});

function asyncHandler(
  fn: (req: Request, res: Response) => Promise<void>
): (req: Request, res: Response, next: NextFunction) => void {
  return (req, res, next) => {
    fn(req, res).catch(next);
  };
}

function sessionSummary(s: Session) {
  return {
    id: s.id,
    originalName: s.originalName,
    status: s.status,
    stages: s.stages,
    error: s.errorMessage,
    settings: {
      paperType: s.settings.paperType,
      mode: s.settings.mode,
      preserveWording: s.settings.preserveWording,
      fixCitationMechanics: s.settings.fixCitationMechanics,
      verifyMetadata: s.settings.verifyMetadata,
    },
  };
}

export function apiRouter(): Router {
  const router = Router();

  router.post(
    "/merge-preview",
    (req, res, next) => {
      mergeUpload.array("documents", 30)(req, res, (err: unknown) => {
        if (err && typeof err === "object" && (err as { code?: string }).code === "LIMIT_FILE_SIZE") {
          next(Errors.tooLarge(config.maxUploadBytes));
        } else if (err) {
          next(Errors.invalid("Upload failed. Select no more than 30 DOCX files."));
        } else next();
      });
    },
    asyncHandler(async (req, res) => {
      const files = (req.files as Express.Multer.File[] | undefined) ?? [];
      if (files.length < 1 || files.length > 30) throw Errors.invalid("Select between 1 and 30 DOCX files.");
      const counts = [];
      for (const file of files) {
        if (path.extname(file.originalname).toLowerCase() !== ".docx") throw Errors.unsupportedType();
        counts.push(await countDocumentWords(file.buffer));
      }
      res.json({ contentWords: counts, appendixSourceWords: await appendixSourceWordCount() });
    })
  );

  router.post(
    "/merge-documents",
    (req, res, next) => {
      mergeUpload.array("documents", 30)(req, res, (err: unknown) => {
        if (err && typeof err === "object" && (err as { code?: string }).code === "LIMIT_FILE_SIZE") {
          next(Errors.tooLarge(config.maxUploadBytes));
        } else if (err) {
          next(Errors.invalid("Upload failed. Select no more than 30 DOCX files."));
        } else next();
      });
    },
    asyncHandler(async (req, res) => {
      const files = (req.files as Express.Multer.File[] | undefined) ?? [];
      if (files.length < 2 || files.length > 30) {
        throw Errors.invalid("Select between 2 and 30 Microsoft Word .docx files.");
      }
      let names: unknown;
      try { names = JSON.parse(String(req.body?.names ?? "[]")); } catch { names = []; }
      if (!Array.isArray(names) || names.length !== files.length) {
        throw Errors.invalid("Every document must have a heading name.");
      }

      const inputs = [];
      for (let index = 0; index < files.length; index++) {
        const file = files[index]!;
        if (path.extname(file.originalname).toLowerCase() !== ".docx") throw Errors.unsupportedType();
        await DocxPackage.load(file.buffer);
        const name = String(names[index] ?? "").trim().slice(0, 200);
        if (!name) throw Errors.invalid("Document headings cannot be empty.");
        inputs.push({ name, originalName: file.originalname, buffer: file.buffer });
      }

      const appendixWords = Number.parseInt(String(req.body?.appendixWords ?? "0"), 10);
      if (!Number.isInteger(appendixWords) || appendixWords < 0 || appendixWords > 50_000) {
        throw Errors.invalid("Appendix word count must be between 0 and 50,000 words.");
      }
      const output = await mergeDocuments(inputs, appendixWords);
      if (output.length < 5 || output.subarray(0, 5).toString("latin1") !== "%PDF-") {
        throw Errors.internal();
      }
      log.info("documents merged", { files: inputs.length, bytes: output.length, referencesRemoved: true, appendixWords });
      res
        .setHeader("Content-Type", "application/pdf")
        .setHeader("Content-Disposition", 'attachment; filename="Merged_Submissions.pdf"')
        .send(output);
    })
  );

  // --- Upload ----------------------------------------------------------
  router.post(
    "/documents",
    (req, res, next) => {
      upload.single("file")(req, res, (err: unknown) => {
        if (err && typeof err === "object" && (err as { code?: string }).code === "LIMIT_FILE_SIZE") {
          next(Errors.tooLarge(config.maxUploadBytes));
        } else if (err) {
          next(Errors.invalid("Upload failed."));
        } else next();
      });
    },
    asyncHandler(async (req, res) => {
      const file = req.file;
      if (!file) throw Errors.invalid("No file was uploaded.");
      const name = file.originalname || "document.docx";
      const ext = path.extname(name).toLowerCase();
      if (ext !== ".docx") throw Errors.unsupportedType();
      const mimeOk = [
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/octet-stream",
        "application/zip",
      ].includes(file.mimetype);
      if (!mimeOk) throw Errors.unsupportedType();

      // Deep validation: real ZIP, safe package, not macro-enabled.
      const pkg = await DocxPackage.load(file.buffer);
      const model = await buildDocumentModel(pkg);
      const analysis = analyzeDocument(model);

      const session = await createSession(name, file.buffer, defaultSettings());
      session.cachedAnalysis = analysis;
      await saveSession(session);

      log.info("document uploaded", {
        sessionId: session.id,
        bytes: file.buffer.length,
        paragraphs: model.paragraphs.length,
        tables: model.tables.length,
        images: model.imageCount,
      });

      res.status(201).json({
        ...sessionSummary(session),
        detected: {
          metadata: analysis.detectedMetadata,
          hasTitlePage: analysis.hasTitlePage,
          hasAbstract: analysis.abstractHeadingIndex != null,
          headings: analysis.headings.length,
          citations: analysis.citations.length,
          references: analysis.references.length,
          tables: model.tables.length,
          images: model.imageCount,
          paragraphs: model.paragraphs.length,
          footnotes: model.footnoteCount,
        },
      });
    })
  );

  // --- Configure + start processing ------------------------------------
  router.post(
    "/documents/:id/process",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (session.status === "processing") {
        throw Errors.notReady("The document is already being processed.");
      }
      const parsed = SettingsSchema.safeParse(req.body ?? {});
      if (!parsed.success) {
        throw Errors.invalid("Invalid settings payload.");
      }
      const s = parsed.data;
      session.settings = {
        paperType: s.paperType,
        mode: s.mode,
        preserveWording: s.preserveWording,
        fixCitationMechanics: s.fixCitationMechanics,
        verifyMetadata: s.verifyMetadata,
        metadata: s.metadata,
        instructor: parseInstructorRequirements(s.instructorRequirements),
      };
      await saveSession(session);
      // Fire and monitor via /status. Errors are captured on the session.
      runInBackground(processSession(session));
      res.status(202).json(sessionSummary(session));
    })
  );

  // --- Status ----------------------------------------------------------
  router.get(
    "/documents/:id/status",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      res.json(sessionSummary(session));
    })
  );

  // --- Report ----------------------------------------------------------
  router.get(
    "/documents/:id/report",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (session.status !== "ready" || !session.report) {
        throw Errors.notReady("The report is not ready yet.");
      }
      res.json({
        ...sessionSummary(session),
        report: session.report,
        outline: buildOutline(session),
        instructorUninterpreted: session.settings.instructor.uninterpreted,
      });
    })
  );

  // --- HTML report download -------------------------------------------
  router.get(
    "/documents/:id/report.html",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (!session.report) throw Errors.notReady("The report is not ready yet.");
      const html = renderReportHtml(session.report, session.originalName);
      res
        .setHeader("Content-Type", "text/html; charset=utf-8")
        .setHeader(
          "Content-Disposition",
          `attachment; filename="${baseName(session)}_APA7_report.html"`
        )
        .send(html);
    })
  );

  // --- Resolve an issue -------------------------------------------------
  router.post(
    "/documents/:id/resolve",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (!session.report) throw Errors.notReady("Process the document first.");
      const parsed = ResolveSchema.safeParse(req.body ?? {});
      if (!parsed.success) throw Errors.invalid("Invalid resolution payload.");
      applyResolution(session, parsed.data.issueKey, parsed.data.optionId, parsed.data.note);
      await saveSession(session);
      res.json({ report: session.report });
    })
  );

  // --- Exclude rules and regenerate ------------------------------------
  router.post(
    "/documents/:id/rules",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      const parsed = RulesSchema.safeParse(req.body ?? {});
      if (!parsed.success) throw Errors.invalid("Invalid rules payload.");
      session.disabledRules = new Set(parsed.data.disabledRules);
      await saveSession(session);
      res.json({ disabledRules: [...session.disabledRules] });
    })
  );

  // --- Regenerate (reprocess from the pristine original) ---------------
  router.post(
    "/documents/:id/generate",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (session.status === "processing") {
        throw Errors.notReady("The document is already being processed.");
      }
      runInBackground(processSession(session));
      res.status(202).json(sessionSummary(session));
    })
  );

  // --- Download corrected document -------------------------------------
  router.get(
    "/documents/:id/download",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      if (session.status !== "ready" || !session.outputPath) {
        throw Errors.notReady("The corrected document is not ready yet.");
      }
      const buffer = await readOutput(session);
      const suffix =
        session.settings.mode === "format_verify" ? "APA7_verified" : "APA7_formatted";
      res
        .setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        .setHeader(
          "Content-Disposition",
          `attachment; filename="${baseName(session)}_${suffix}.docx"`
        )
        .send(buffer);
    })
  );

  // --- Download pristine original --------------------------------------
  router.get(
    "/documents/:id/original",
    asyncHandler(async (req, res) => {
      const session = await getSession(req.params.id!);
      const buffer = await readOriginal(session);
      res
        .setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        .setHeader(
          "Content-Disposition",
          `attachment; filename="${session.originalName}"`
        )
        .send(buffer);
    })
  );

  // --- Delete now -------------------------------------------------------
  router.delete(
    "/documents/:id",
    asyncHandler(async (req, res) => {
      await deleteSession(req.params.id!);
      res.status(204).end();
    })
  );

  router.get("/health", (_req, res) => {
    res.json({ ok: true, product: config.productName });
  });

  return router;
}

function baseName(session: Session): string {
  return session.originalName.replace(/\.docx$/i, "");
}

/** Lightweight structural outline for the document viewer. */
function buildOutline(session: Session) {
  const a = session.cachedAnalysis;
  if (!a) return [];
  const headingByIndex = new Map(a.headings.map((h) => [h.paragraphIndex, h]));
  const refSet = new Set(a.referenceEntryIndexes);
  const issuesByParagraph = new Map<number, number>();
  for (const issue of session.report?.issues ?? []) {
    const idx = issue.location?.paragraphIndex;
    if (idx != null) issuesByParagraph.set(idx, (issuesByParagraph.get(idx) ?? 0) + 1);
  }
  const out: {
    index: number;
    kind: string;
    level?: number;
    confidence?: string;
    text: string;
    issues: number;
  }[] = [];
  // The analysis retains only indexes/text via citations etc.; rebuild from
  // heading and reference info plus title metadata.
  const titleIdx = a.detectedMetadata.titleParagraphIndex;
  const push = (index: number, kind: string, text: string, extra?: { level?: number; confidence?: string }) => {
    out.push({
      index,
      kind,
      text: text.length > 240 ? text.slice(0, 239) + "…" : text,
      issues: issuesByParagraph.get(index) ?? 0,
      ...extra,
    });
  };
  // We stored paragraph text in citations/references only; use analysis via headings/references and title.
  if (titleIdx != null) push(titleIdx, "title", a.detectedMetadata.title ?? "");
  if (a.abstractHeadingIndex != null) push(a.abstractHeadingIndex, "abstract-heading", "Abstract");
  for (const h of a.headings) {
    push(h.paragraphIndex, "heading", h.text, {
      level: h.level,
      confidence: h.confidence,
    });
  }
  if (a.referencesHeadingIndex != null) {
    push(a.referencesHeadingIndex, "references-heading", "References");
  }
  for (const r of a.references) {
    push(r.paragraphIndex, "reference", r.raw);
  }
  return out.sort((x, y) => x.index - y.index);
}

/** Central error handler — never leaks stack traces. */
export function errorHandler(
  err: unknown,
  _req: Request,
  res: Response,
  _next: NextFunction
): void {
  if (err instanceof AppError) {
    res.status(err.httpStatus).json({
      error: { code: err.code, message: err.userMessage },
    });
    return;
  }
  log.error("unhandled error", {
    error: err instanceof Error ? err.message : String(err),
  });
  res.status(500).json({
    error: { code: "INTERNAL", message: Errors.internal().userMessage },
  });
}
