# Claim-level citation verification

## Configuration

Set `OPENAI_API_KEY` and `OPENAI_CITATION_MODEL` in the server environment (Vercel project Environment Variables for production), then redeploy. Choose an OpenAI Responses API model supporting strict structured outputs. Never use a `VITE_` prefix or put secrets in the repository. Calls use `store: false`.

The upload screen discloses that document passages and source text go to OpenAI. Existing syntax-only document correction remains available without credentials. A configured key enables paid requests from this public website; configure account spend limits as appropriate for the site's traffic.

## Implemented workflow

1. DOCX upload creates a random session and separate access token. The original is retained; the existing citation-mechanics pass creates the analysis baseline.
2. Reference DOI metadata is retrieved from Crossref. Available abstracts are used as abstracts, not full-text verification. Otherwise the reference URL is fetched. Redirects and DNS destinations are validated and pinned to public IPv4 addresses. HTML and text PDFs are readable; scanned text is not OCRed.
3. Every nonempty document-body paragraph before References (including table paragraphs) goes through a claim assessment. Nearby paragraphs provide context. Sentence coverage gaps are reported for review.
4. Proposed edits receive a second, separate model assessment before the deterministic edit gate checks original-text anchors, source IDs, source access status, exact evidence quotations, citation metadata, numeric values, and edit overlap.
5. Only supported insertions and isolated parenthetical replacements pass the gate. Citation strings come from parsed references, never model-invented author/year data. Edits operate on Word text nodes and preserve surrounding runs/properties/package parts.
6. The output package is reopened and checked for integrity and explained content changes. Citation/reference matching is rerun. Users download DOCX and an escaped, self-contained HTML verification report.

Jobs use existing KV/session persistence, token-gated status/step/download endpoints and a versioned lock. Each step is one source or model request, fitting the current 60-second function limit. Refresh preserves the current job token in sessionStorage; Resume continues at the saved checkpoint. Failed steps do not advance the saved version. Session retention still applies.

## Explicit limits / outstanding acceptance work

This is an AI-assisted assessment, not a guarantee of scholarly correctness. Exact quotation checks prove that the quotation was retrieved; semantic support is still model-assessed.

- Without OpenAI credentials the live semantic evaluation cannot run. Simulated-provider tests establish transport/edit correctness only.
- Retrieval is limited to DOI abstracts and the provided reference URL. No paywall bypass, library subscription access, OCR or automatic open-access mirror discovery. A DOI metadata year conflict requires review.
- Each source response is limited to 2 MB. Source excerpts are limited to 16,000 characters; PDFs to at most 20 pages within that limit. Reports disclose this scope. Aggregate source context above 260,000 characters stops with a clear error; paragraphs above 12,000 characters are reported as unreviewed.
- Headers/footers, footnotes/endnotes and image-only text are not claim-audited. Document-body text extraction follows the existing Word model.
- Narrative source replacement, automatic citation relocation/removal, unresolved same-author/year disambiguation, secondary-source reconstruction, and direct-quotation locator completion remain manual review. They are not claimed as completed acceptance criteria.
- It cannot know which deleted citation originally accompanied a rewritten sentence. It can propose a supported attribution only from retrieved evidence.

## Validation

Run `npm test --workspace server` and `npm run build`.

The `citation_evidence` and `citation_evidence_api` tests use explicitly synthetic evidence/model outputs to test the citation gates, formatting preservation, token checks, retries, idempotency and both downloads.

The opt-in `citation_evidence_live.test.ts` uses real Barr/Gibbs, Dahl/Lochner, Carneiro/Ginja and Sowell references with real retrieval and paid OpenAI inference. Set `RUN_LIVE_CITATION_EVAL=1` alongside the API key/model, then run `npm test --workspace server -- citation_evidence_live`. It fails if a required source cannot be accessed, rather than counting missing evidence as a successful semantic evaluation. Do not claim full acceptance until this and additional real-document evaluations pass.

API contract: https://developers.openai.com/api/docs/guides/structured-outputs
APA guidance: https://apastyle.apa.org/style-grammar-guidelines/citations/basic-principles/author-date
