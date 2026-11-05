import { createHash, randomBytes, randomUUID } from "node:crypto";
import { kv } from "@vercel/kv";
import { DocxPackage, verifyDocxIntegrity } from "../../../docx/package.js";
import { buildDocumentModel, contentFingerprint } from "../../../docx/model.js";
import { rawParagraphText, replaceParagraphTextAt } from "../../../docx/text.js";
import { descendantsW } from "../../../docx/xml.js";
import { assertContentPreserved } from "../../../pipeline.js";
import { createSession, getSession, readOriginal, writeOutput, saveSession } from "../../../store/sessions.js";
import { AppError, Errors } from "../../../errors.js";
import { defaultSettings } from "../../requirements.js";
import { analyzeDocument } from "../../analysis.js";
import type { Change } from "../../types.js";
import { correctCitationDocument } from "../correct_document.js";
import { retrieveSource } from "./sources.js";
import { assessParagraph, evidenceConfigured } from "./openai.js";
import { planEdits } from "./edits.js";
import type { EvidenceJob } from "./types.js";

const hash = (s: string) => createHash("sha256").update(s).digest("hex");
export async function startEvidenceJob(name: string, input: Buffer) {
  if (!evidenceConfigured()) throw new AppError("VERIFICATION_UNAVAILABLE", "Source verification needs OPENAI_API_KEY and OPENAI_CITATION_MODEL in the server environment.", 503);
  const corrected = await correctCitationDocument(input);
  const model = await buildDocumentModel(await DocxPackage.load(corrected.output));
  const analysis = analyzeDocument(model);
  const token = randomBytes(32).toString("hex");
  const job: EvidenceJob = { tokenHash: hash(token), version: 0, phase: "sources", sourceCursor: 0, paragraphCursor: 0,
    references: analysis.references, sources: [], claims: [], edits: [], mechanics: corrected.summary.changes,
    warnings: [...corrected.summary.warnings], uncitedReferences: corrected.summary.uncitedReferences,
    model: process.env.OPENAI_CITATION_MODEL!,
    paragraphs: model.paragraphs.filter((p) => !p.isEmpty && (analysis.referencesHeadingIndex == null || p.index < analysis.referencesHeadingIndex))
      .map((p) => ({ index: p.index, text: rawParagraphText(p.el), editable: !["fldChar", "instrText", "ins", "del"].some((tag) => descendantsW(p.el, tag).length > 0) })),
  };
  if (!job.references.length) job.warnings.push("No references detected; claims can be reviewed but no sources can be assigned.");
  const session = await createSession(name, input, { ...defaultSettings(), mode: "format", verifyMetadata: false });
  session.citationJob = job; session.status = "processing"; await saveSession(session);
  return { id: session.id, token, ...jobStatus(job) };
}
export function jobStatus(job: EvidenceJob) {
  return { phase: job.phase, version: job.version, sourcesTotal: job.references.length, sourcesDone: job.sourceCursor,
    paragraphsTotal: job.paragraphs.length, paragraphsDone: job.paragraphCursor,
    applied: job.claims.filter((c) => c.applied).length, review: job.claims.filter((c) => c.review || ["partial", "insufficient", "contradicted"].includes(c.status)).length };
}
export async function authorizedJob(id: string, token: string) {
  const session = await getSession(id, true);
  if (!session.citationJob || hash(token) !== session.citationJob.tokenHash) throw Errors.notFound();
  return session;
}
const localLocks = new Set<string>();
const useKv = () => Boolean(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);
export async function stepEvidenceJob(id: string, token: string, version: number) {
  // Atomic lease across serverless invocations; duplicate browser requests do not reapply edits.
  const lock = `apa7:citation-lock:${id}`, owner = randomUUID();
  if (useKv() ? !await kv.set(lock, owner, { nx: true, ex: 90 }) : localLocks.has(id)) throw Errors.notReady("This document is already processing a step.");
  localLocks.add(id);
  try {
    const session = await authorizedJob(id, token);
    // Work on a copy so failures cannot mutate the in-memory saved checkpoint.
    const job = structuredClone(session.citationJob!);
    if (version !== job.version || job.phase === "complete") return jobStatus(job);
    if (job.phase === "sources") {
      if (job.sourceCursor < job.references.length) job.sources.push(await retrieveSource(job.references[job.sourceCursor]!, job.sourceCursor++));
      if (job.sourceCursor >= job.references.length) job.phase = "claims";
    } else if (job.phase === "claims") {
      const p = job.paragraphs[job.paragraphCursor];
      if (p) {
        if (p.text.length > 12000) job.warnings.push(`Paragraph ${p.index + 1} exceeds the claim review limit; review manually.`);
        else {
          const context = job.paragraphs.slice(Math.max(0, job.paragraphCursor - 1), job.paragraphCursor + 2).map((q) => q.text).join("\n");
          if (!job.pending) {
            job.pending = await assessParagraph(p.text, job.sources, context, job.model);
            job.version++; session.citationJob = job; await saveSession(session); return jobStatus(job);
          }
          const decisions = job.pending.some((d) => d.action !== "none")
            ? await assessParagraph(p.text, job.sources, context, job.model, job.pending) : job.pending;
          delete job.pending;
          const planned = planEdits(p.text, p.index, decisions, job.sources, p.editable);
          job.claims.push(...planned.claims); job.edits.push(...planned.edits);
          if (!decisions.length && p.text.split(/\s+/).length >= 12) job.warnings.push(`Paragraph ${p.index + 1}: model returned no claim decisions; review manually.`);
        }
        job.paragraphCursor++;
      }
      if (job.paragraphCursor >= job.paragraphs.length) job.phase = "finalize";
    } else {
      const mechanical = await correctCitationDocument(await readOriginal(session));
      const pkg = await DocxPackage.load(mechanical.output), model = await buildDocumentModel(pkg);
      const before = contentFingerprint(model), changes: Change[] = [];
      for (const p of model.paragraphs) {
        const edits = job.edits.filter((e) => e.paragraph === p.index).sort((a, b) => b.start - a.start);
        if (!edits.length) continue;
        const original = rawParagraphText(p.el);
        if (original !== job.paragraphs.find((q) => q.index === p.index)?.text) throw Errors.invalid("Document text changed during verification.");
        for (const e of edits) if (!replaceParagraphTextAt(p.el, e.start, e.length, e.value)) throw Errors.invalid("Citation could not be applied safely.");
        changes.push({ id: randomUUID(), ruleId: "APA-CITATION-EVIDENCE", category: "citations", location: { paragraphIndex: p.index },
          before: original, after: rawParagraphText(p.el), reason: "Apply evidence-backed claim citation.", confidence: 0.95, stage: "format", timestamp: new Date().toISOString() });
      }
      if (changes.length) pkg.markDirty("word/document.xml");
      const output = changes.length ? await pkg.save() : mechanical.output;
      await verifyDocxIntegrity(output);
      const out = await buildDocumentModel(await DocxPackage.load(output));
      assertContentPreserved(before, contentFingerprint(out), changes);
      const audited = analyzeDocument(out);
      const used = new Set(audited.matching.citationMatches.flatMap((m) => m.referenceIndexes));
      job.uncitedReferences = audited.references.filter((_r, i) => !used.has(i)).map((r) => r.raw);
      await writeOutput(session, output); session.status = "ready"; job.phase = "complete";
    }
    job.version++; session.citationJob = job; await saveSession(session); return jobStatus(job);
  } finally {
    localLocks.delete(id);
    if (useKv()) await kv.eval("if redis.call('get',KEYS[1]) == ARGV[1] then return redis.call('del',KEYS[1]) else return 0 end", [lock], [owner]);
  }
}

const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
export function evidenceReport(job: EvidenceJob): string {
  return `<!doctype html><html lang="en"><meta charset="utf-8"><title>Citation verification report</title><style>body{font:16px/1.6 system-ui;max-width:1000px;margin:40px auto;padding:20px}article{border:1px solid #ccc;padding:20px;margin:20px 0}blockquote{border-left:3px solid #aaa;padding:10px}pre{white-space:pre-wrap}</style><h1>Citation verification report</h1><p>AI-assisted source assessment (${escape(job.model)}). Decisions are limited to accessible source text; human review remains appropriate. Exact evidence quotes were checked against retrieved content before applying citations.</p>
    <h2>Mechanical corrections</h2>${job.mechanics.map((m) => `<p>Paragraph ${m.paragraph}: ${escape(m.before)} → ${escape(m.after)}</p>`).join("")}
    <h2>Claim decisions</h2>${job.claims.map((c) => `<article><h3>Paragraph ${c.paragraph + 1}: ${escape(c.status)} — ${c.applied ? "correction applied" : "no edit"}</h3><p>${escape(c.sentence)}</p><p>Claim: ${escape(c.claim)}</p><p>Original citation: ${escape(c.originalCitation || "None specified")}</p><p>Corrected citation: ${escape(c.correctedCitation || "None")}</p><p>Action: ${escape(c.action)}; ${escape(c.reason)}</p><p>${escape(c.review)}</p>${c.evidence.map((e) => { const s = job.sources.find((s) => s.id === e.sourceId); return `<p>${escape(s?.reference.raw ?? "Unknown source")}</p><p>${escape(s?.url ?? "")}</p><blockquote>${escape(e.quote)}</blockquote>`; }).join("")}</article>`).join("")}
    <h2>Source access and scope</h2>${job.sources.map((s) => `<p>${escape(s.reference.raw)}<br>${escape(s.status)}: ${escape(s.scope)}<br>${escape(s.url)}</p>`).join("")}
    <h2>Uncited references</h2><ul>${job.uncitedReferences.map((r) => `<li>${escape(r)}</li>`).join("")}</ul><h2>Items requiring review</h2><ul>${job.warnings.map((w) => `<li>${escape(w)}</li>`).join("")}</ul></html>`;
}
