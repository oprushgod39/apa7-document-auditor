import path from "node:path";
import { Router, type NextFunction, type Request, type Response } from "express";
import multer from "multer";
import { z } from "zod";
import { Errors } from "../errors.js";
import {
  fetchReport,
  getAvailability,
  getDocument,
  isSimilarityScanConfigured,
  reportUrl,
  unlockAi,
  uploadDocument,
  validPercentage,
  type SimilarityScanDocument,
} from "./client.js";
import { authorizeSimilarityScanUpload, readEncryptedUpload, removeEncryptedUpload, SIMILARITYSCAN_MAX_UPLOAD_BYTES } from "./blob_transport.js";
import {
  createSimilarityScanSession,
  deleteSimilarityScanSession,
  getSimilarityScanSession,
  readSimilarityScanOriginal,
  verifySimilarityScanAccess,
  type SimilarityScanSession,
} from "./store.js";

const SUPPORTED_EXTENSIONS = new Set([".pdf", ".docx", ".txt", ".rtf", ".xlsx"]);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: SIMILARITYSCAN_MAX_UPLOAD_BYTES, files: 1 } });
const BlobUploadSchema = z.object({
  batchId: z.string().uuid(),
  url: z.string().url().max(2048),
  encryptionKey: z.string().min(43).max(64),
  iv: z.string().min(16).max(32),
  filename: z.string().trim().min(1).max(240),
  size: z.number().int().positive().max(SIMILARITYSCAN_MAX_UPLOAD_BYTES),
  contentType: z.string().max(200).default("application/octet-stream"),
});

function asyncHandler(fn: (req: Request, res: Response) => Promise<void>) {
  return (req: Request, res: Response, next: NextFunction) => { fn(req, res).catch(next); };
}

function validateFilename(filename: string): void {
  if (!SUPPORTED_EXTENSIONS.has(path.extname(filename).toLowerCase())) {
    throw Errors.invalid("Choose a PDF, DOCX, TXT, RTF, or XLSX file.");
  }
}

function accessToken(req: Request): string | undefined {
  return req.header("X-ScholarlyWorks-Session")?.trim() || undefined;
}

async function authorized(req: Request): Promise<SimilarityScanSession> {
  const session = await getSimilarityScanSession(req.params.id!);
  verifySimilarityScanAccess(session, accessToken(req));
  return session;
}

function summary(session: SimilarityScanSession, document: SimilarityScanDocument) {
  const aiLocked = document.ai?.locked === true;
  const terminal = ["completed", "error", "failed_invalid", "cancelled"].includes(document.status);
  const progress = document.status === "completed" ? 100
    : document.status === "processing" ? 65
      : document.status === "queued" ? 35
        : document.status === "submitted" ? 20 : 0;
  const apiMessage = typeof document.error === "string" ? document.error : document.error?.message;
  return {
    id: session.id,
    orderId: session.orderId,
    filename: document.fileName || session.originalName,
    status: document.status,
    terminal,
    progress,
    similarityPercentage: validPercentage(document.similarity?.percentage),
    aiPercentage: validPercentage(document.ai?.percentage),
    similarityReportAvailable: Boolean(reportUrl(document, "similarity")),
    aiReportAvailable: Boolean(reportUrl(document, "ai")),
    aiLocked,
    aiUnlockPrice: aiLocked && typeof document.ai?.unlockPrice === "number" ? document.ai.unlockPrice : null,
    message: document.message || apiMessage || null,
  };
}

export function similarityScanRouter(): Router {
  const router = Router();

  router.get("/availability", asyncHandler(async (_req, res) => {
    if (!isSimilarityScanConfigured()) {
      res.json({ configured: false, ready: false, acceptingUploads: false, status: "unavailable", message: "Checker credentials are not configured.", maxUploadBytes: SIMILARITYSCAN_MAX_UPLOAD_BYTES });
      return;
    }
    const status = await getAvailability();
    res.json({ configured: true, ...status, maxUploadBytes: SIMILARITYSCAN_MAX_UPLOAD_BYTES });
  }));

  router.post("/upload", asyncHandler(async (req, res) => {
    const availability = await getAvailability();
    if (!availability.ready || !availability.acceptingUploads) {
      throw Errors.notReady(availability.message || "The checking service is not accepting uploads right now.");
    }
    res.json(await authorizeSimilarityScanUpload(req, req.body));
  }));

  router.post(
    "/documents",
    (req, res, next) => {
      if (req.is("application/json")) return next();
      upload.single("file")(req, res, (error: unknown) => {
        if (error && typeof error === "object" && (error as { code?: string }).code === "LIMIT_FILE_SIZE") {
          next(Errors.tooLarge(SIMILARITYSCAN_MAX_UPLOAD_BYTES));
        } else if (error) next(Errors.invalid("The file could not be uploaded."));
        else next();
      });
    },
    asyncHandler(async (req, res) => {
      let contents: Buffer;
      let input: {
        filename: string;
        size: number;
        contentType: string;
        blob?: { url: string; batchId: string; encryptionKey: string; iv: string };
      };
      if (req.is("application/json")) {
        const parsed = BlobUploadSchema.safeParse(req.body);
        if (!parsed.success) throw Errors.invalid("Invalid secure checker upload request.");
        const value = parsed.data;
        const blob = { url: value.url, batchId: value.batchId, encryptionKey: value.encryptionKey, iv: value.iv };
        input = {
          filename: value.filename,
          size: value.size,
          contentType: value.contentType,
          blob,
        };
        contents = await readEncryptedUpload({ ...blob, expectedBytes: value.size });
      } else {
        const file = req.file;
        if (!file) throw Errors.invalid("Choose a file to check.");
        input = { filename: file.originalname, size: file.size, contentType: file.mimetype };
        contents = file.buffer;
      }
      validateFilename(input.filename);
      let uploaded;
      try {
        uploaded = await uploadDocument({ contents, filename: input.filename, contentType: input.contentType });
      } catch (error) {
        if (input.blob) {
          try { await removeEncryptedUpload(input.blob.url, input.blob.batchId); } catch { /* best effort */ }
        }
        throw error;
      }
      const created = await createSimilarityScanSession({
        orderId: uploaded.orderId,
        originalName: input.filename,
        originalSize: input.size,
        originalContentType: input.contentType,
        ...(input.blob ? { blob: input.blob } : { buffer: contents }),
      });
      const document: SimilarityScanDocument = {
        success: true,
        orderId: uploaded.orderId,
        status: uploaded.status,
        fileName: created.session.originalName,
      };
      res.status(202).json({ ...summary(created.session, document), accessToken: created.accessToken });
    })
  );

  router.get("/documents/:id/status", asyncHandler(async (req, res) => {
    const session = await authorized(req);
    res.json(summary(session, await getDocument(session.orderId)));
  }));

  router.post("/documents/:id/unlock-ai", asyncHandler(async (req, res) => {
    const session = await authorized(req);
    const before = await getDocument(session.orderId);
    if (before.status !== "completed" || before.ai?.locked !== true) {
      throw Errors.invalid("The AI report is not locked for this completed document.");
    }
    await unlockAi(session.orderId);
    res.json(summary(session, await getDocument(session.orderId)));
  }));

  router.get("/documents/:id/report/:kind", asyncHandler(async (req, res) => {
    const session = await authorized(req);
    const kind = req.params.kind;
    if (kind !== "similarity" && kind !== "ai") throw Errors.notFound();
    const fresh = await getDocument(session.orderId);
    const url = reportUrl(fresh, kind);
    if (!url) {
      throw Errors.notReady(kind === "ai" ? "AI report not available." : "Similarity report not available.");
    }
    const report = await fetchReport(url);
    res.setHeader("Content-Type", report.contentType);
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("Content-Disposition", `inline; filename="${kind === "ai" ? "AI_Report" : "Similarity_Report"}.pdf"`);
    res.send(report.contents);
  }));

  router.get("/documents/:id/original", asyncHandler(async (req, res) => {
    const session = await authorized(req);
    res.setHeader("Content-Type", session.originalContentType);
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("Content-Disposition", `attachment; filename="${session.originalName.replace(/["\r\n]/g, "_")}"`);
    res.send(await readSimilarityScanOriginal(session));
  }));

  router.delete("/documents/:id", asyncHandler(async (req, res) => {
    const session = await authorized(req);
    await deleteSimilarityScanSession(session.id);
    res.status(204).end();
  }));

  return router;
}
