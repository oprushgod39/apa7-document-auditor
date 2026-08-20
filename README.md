# APA 7 Document Auditor

Upload a Microsoft Word `.docx` academic paper, analyze it against APA 7th
Edition requirements, automatically apply safe deterministic corrections,
audit citations and references, verify reference metadata against Crossref,
resolve uncertain issues interactively, and download a corrected `.docx` plus
a detailed compliance report.

**Product principles**

- The paper's wording is **never** rewritten, paraphrased, or shortened
  ("Preserve wording — ON" by default). Only formatting, layout, and
  deterministic citation/reference mechanics are touched.
- No fake certainty: every rule ends in one of
  `PASS / FIXED / WARNING / USER REVIEW / FAIL / NOT APPLICABLE / UNVERIFIED`.
  **APA Validated** means *every applicable rule has either been automatically
  verified or explicitly resolved by the user* — never a guarantee that an
  arbitrary paper is "100% APA".
- The original upload is preserved untouched and every automatic change is
  recorded (rule ID, location, before/after, reason, confidence).
- An **independent audit pass** re-checks the *modified* document; a rule is
  never assumed correct just because a fix was applied.
- A **content-preservation guard** compares semantic fingerprints
  (paragraph text, images, table shapes, hyperlinks, footnotes, equations)
  before/after formatting and aborts the run rather than ship an unexpected
  content change.

The product name is configurable: `PRODUCT_NAME` env (server/logs/report) and
the `PRODUCT_NAME` constant in `web/src/App.tsx` (UI).

---

## Architecture

```
apa7-document-auditor/
├── server/                    Node 20+ · Express · TypeScript (ESM)
│   ├── src/
│   │   ├── docx/              OOXML engine
│   │   │   ├── package.ts     Safe ZIP load/save, macro & ZIP-bomb rejection,
│   │   │   │                  post-generation corruption check
│   │   │   ├── model.ts       Document model + style-chain resolution
│   │   │   ├── edit.ts        Schema-order-aware OOXML mutations
│   │   │   ├── text.ts        Cross-run-safe text replacement
│   │   │   └── xml.ts         Namespaced DOM helpers (DOCTYPE rejected)
│   │   ├── apa/
│   │   │   ├── types.ts       Rule / Issue / Change domain models
│   │   │   ├── engine.ts      Rule registry + execution
│   │   │   ├── analysis.ts    One-pass structural analysis (cached)
│   │   │   ├── requirements.ts  APA baseline ⊕ instructor overrides
│   │   │   ├── rules/         layout, paragraphs, title_page, headings,
│   │   │   │                  abstract, citations, references, quotations,
│   │   │   │                  tables, figures, statistics, numbers, abbreviations
│   │   │   ├── citations/     Deterministic in-text citation parser
│   │   │   ├── references/    Reference parser, type classifier, bidirectional
│   │   │   │                  citation ↔ reference matcher
│   │   │   └── headings/      Multi-signal heading classifier (scored)
│   │   ├── verify/            Metadata provider abstraction + Crossref impl
│   │   ├── audit/             Independent auditor + HTML report renderer
│   │   ├── store/             Session store: in-memory Map + local temp files
│   │   │                      (default), or Vercel KV + Blob when configured
│   │   ├── pipeline.ts        analyze → format → guard → verify → audit
│   │   └── api/routes.ts      REST API (zod-validated)
│   └── test/                  vitest: unit, golden-document, round-trip,
│                              idempotency, security, API integration
├── web/                       React 18 · Vite · TypeScript SPA
└── Dockerfile / docker-compose.yml
```

**Why Node for DOCX?** The engine treats the `.docx` as an OOXML ZIP and
mutates only the XML nodes a rule needs (JSZip + XML DOM). Untouched parts —
images, equations, footnotes, themes, custom XML, embedded media — are carried
through byte-identical, which is stronger preservation than a
parse-everything/rebuild-everything library. There is **no** DOCX → HTML →
DOCX round trip anywhere.

### Processing pipeline

```
original.docx (never modified)
  → safe package load (macro/ZIP-bomb/XXE checks)
  → document model + structural analysis (cached per session)
  → rule engine, fix mode (HIGH-confidence deterministic fixes only)
  → save package → corruption check (reopen, validate all XML)
  → content-preservation guard (abort on unexplained content change)
  → optional Crossref metadata verification (graceful degradation)
  → INDEPENDENT audit: full rule engine, check-only, on the OUTPUT document
  → compliance report (state: APA Validated / Review Required)
  → user resolves USER REVIEW items → optional regeneration from the original
```

---

## Local setup

Requirements: Node.js ≥ 20 and npm.

```bash
npm install
```

### Development (two processes)

```bash
npm run dev        # API server on http://localhost:8000 (tsx watch)
```

```bash
npm run dev:web    # Vite dev server on http://localhost:5173, proxies /api
```

Open http://localhost:5173.

### Production build (single process)

```bash
npm run build
```

```bash
npm start
```

Open http://localhost:8000 — the server serves the built frontend and the API.

### Tests & checks

```bash
npm test           # server test suite (vitest)
```

```bash
npm run typecheck  # strict TypeScript across server and web
```

### Docker

```bash
docker compose up --build
```

Open http://localhost:8000.

### Deployment (Vercel)

The app also deploys as Vercel serverless functions, with no code changes
from the local setup above:

- `api/index.ts` adapts the same Express app (`server/src/app.ts`) to
  Vercel's Node function convention.
- `vercel.json` builds `web/` as a static SPA (`npm run build` at the repo
  root) and rewrites `/api/*` to the function, everything else to
  `index.html`.
- `server/src/store/sessions.ts` auto-detects Vercel KV / Blob env vars at
  runtime (see below) and persists session metadata / document binaries
  there instead of the in-memory Map + local temp files used locally —
  necessary because serverless functions are stateless across invocations
  and only have ephemeral `/tmp`. With neither attached, the deployed
  function still runs using the in-memory/local fallback, but state will
  not survive between separate invocations (e.g. a cold start between
  upload and status-poll requests) — attach at least KV for a working
  deployment.
- File uploads already use multer's in-memory storage, so no change was
  needed for serverless request handling.

To enable it: import the repo into a Vercel project, then in the project's
**Storage** tab create and connect a **KV** database and a **Blob** store
(this auto-populates the env vars below) and redeploy.

---

## Environment variables

See [.env.example](.env.example). Highlights:

| Variable | Default | Purpose |
| --- | --- | --- |
| `PORT` | `8000` | API/frontend port |
| `MAX_UPLOAD_SIZE` | `26214400` | Upload cap in bytes (25 MB) |
| `FILE_RETENTION_MINUTES` | `60` | Automatic deletion of stored files |
| `STORAGE_DIR` | OS temp subdir | Temporary document storage |
| `METADATA_PROVIDER` | `crossref` | `crossref` or `none` |
| `CROSSREF_MAILTO` | — | Contact email for Crossref polite pool |
| `CROSSREF_MAX_REQUESTS_PER_RUN` | `25` | Per-run request budget |
| `RATE_LIMIT_MAX` / `RATE_LIMIT_WINDOW_MS` | 60 / 60s | API rate limit |
| `KV_REST_API_URL` / `KV_REST_API_TOKEN` (or `KV_URL`) | — | Vercel KV — session metadata backend (Vercel-provisioned) |
| `BLOB_READ_WRITE_TOKEN` | — | Vercel Blob — document binary backend (Vercel-provisioned) |

No secrets are required; never commit a `.env`.

---

## API

| Endpoint | Purpose |
| --- | --- |
| `POST /api/documents` | Upload (multipart `file`); returns detected structure & metadata |
| `POST /api/documents/{id}/process` | Configure settings and start processing |
| `GET  /api/documents/{id}/status` | Processing stages (real, not simulated) |
| `GET  /api/documents/{id}/report` | Compliance report + change log + outline |
| `GET  /api/documents/{id}/report.html` | Downloadable HTML compliance report |
| `POST /api/documents/{id}/resolve` | Record a user resolution for an issue |
| `POST /api/documents/{id}/rules` | Disable rule IDs for regeneration |
| `POST /api/documents/{id}/generate` | Reprocess from the pristine original |
| `GET  /api/documents/{id}/download` | Corrected `…_APA7_formatted/verified.docx` |
| `GET  /api/documents/{id}/original` | The untouched original upload |
| `DELETE /api/documents/{id}` | Delete immediately |

Errors are structured (`{ error: { code, message } }`) and never expose stack
traces: `UNSUPPORTED_FILE_TYPE`, `FILE_TOO_LARGE`, `CORRUPT_DOCUMENT`,
`PASSWORD_PROTECTED`, `MACRO_DOCUMENT_REJECTED`, `UNSAFE_PACKAGE`, …

---

## Rule engine — adding a new APA rule

Rules live in `server/src/apa/rules/` and implement the `ApaRule` interface:

```ts
export const myRules: ApaRule[] = [{
  id: "APA-XYZ-001",           // stable ID surfaced in reports
  category: "layout",
  description: "…",
  severity: "warning",
  applies: (ctx) => true,       // applicability (paper type, structure, …)
  run(ctx, fix) {
    // check the model via ctx.model / ctx.analysis / ctx.req;
    // in fix mode apply HIGH-confidence corrections via docx/edit.ts and
    // record them with ctx.addChange(); report uncertainty with
    // ctx.addIssue({ status: "user_review", resolutionOptions: […] });
    return result("APA-XYZ-001", checked, passed, fixedAny, worstStatus);
  },
}];
```

Register the module in `server/src/apa/engine.ts` (`allRules()`), add tests in
`server/test/`. Design rules so they are **idempotent**: running the formatter
on already-correct output must produce no further changes (enforced by the
idempotency test). Rules can be disabled per session via
`POST /api/documents/{id}/rules` — the hook for future institutional or
instructor profiles.

Instructor/assignment requirements are parsed in `apa/requirements.ts` into an
override layer (font, abstract required/forbidden, running head, minimum
references); interpreted overrides are labeled "Instructor override" in the
report, and uninterpreted lines are surfaced verbatim for manual attention.

---

## Privacy

- Documents are stored under randomized names in a dedicated temp directory
  and deleted automatically after `FILE_RETENTION_MINUTES` (default 60), or
  immediately via `DELETE /api/documents/{id}`.
- Document contents, reference text, and personal data are **never logged** —
  logs carry only IDs, counts, durations, and error codes.
- Nothing is used for training. The only external call is the optional
  Crossref metadata lookup (title/author/year of *references*, never body
  text), which can be disabled with `METADATA_PROVIDER=none` or the UI toggle.
- Uploads are validated in depth: extension + MIME + ZIP magic, entry path
  traversal, decompression limits (ZIP-bomb), DOCTYPE/XXE rejection, macro
  (`vbaProject`, `macroEnabled`) rejection, OLE (password-protected) rejection.
- Security headers (helmet + CSP), API rate limiting, zod-validated payloads,
  no arbitrary file execution.

---

## Known limitations (honest edition)

- **Word-native rendering is not simulated.** Locations are reported by
  paragraph, not page number; title "upper half of page" positioning is
  checked only structurally.
- **Level 4/5 inline headings** (run-in headings sharing a paragraph with
  body text) are supported when user-confirmed but are not auto-detected.
- **Italic sub-segments** (e.g. italicizing just a journal title inside a
  reference run, or "Keywords:" label) are flagged, not auto-fixed, when the
  text shares a single run — splitting runs is riskier than advising.
- **Statistics / numbers / abbreviations rules are advisory only** by design;
  they never modify sentences.
- Crossref covers DOIs/journals well; books, webpages, and reports frequently
  come back `UNVERIFIED` — that is reported as *unverified*, never as failed.
- Citation parsing is deterministic (no LLM anywhere); exotic citation formats
  may be reported as unparsed rather than guessed at.
- Sessions are in-memory (restart clears them; the retention policy makes this
  acceptable for V1). The store is isolated behind `store/sessions.ts` so a
  PostgreSQL-backed store can replace it without touching the pipeline.
- Authentication is intentionally absent in V1; the Express app is structured
  so auth middleware can wrap `/api` later.

## Rules that stay dependent on user judgment

- Heading levels below HIGH confidence (reclassify via the UI: Normal / L1–L5)
- Suspicious heading-hierarchy skips (Level 1 → Level 3)
- Citations with no matching reference (may be personal communications)
- References never cited in the body (may be intentional)
- Missing title-page metadata (never invented — you supply it)
- Crossref mismatches (document value vs. verified value — you choose)
- Long quotations embedded inside larger paragraphs
- Tables/figures lacking labels (titles are never invented)
- Abstract required/forbidden conflicts with assignment instructions
