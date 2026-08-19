import { config } from "../config.js";
import type { ParsedReference } from "../apa/references/parser.js";
import type {
  MetadataProvider,
  VerificationResult,
  VerifiedMetadata,
} from "./provider.js";

/**
 * Crossref metadata provider.
 *
 * - DOI present → direct /works/{doi} lookup (authoritative).
 * - Otherwise → bibliographic search, scored against document values.
 * - Responses cached in-memory; failures degrade to "provider_unavailable"
 *   without failing the document.
 */

interface CrossrefWork {
  DOI?: string;
  title?: string[];
  author?: { family?: string; given?: string }[];
  "container-title"?: string[];
  issued?: { "date-parts"?: number[][] };
  volume?: string;
  issue?: string;
  page?: string;
}

const cache = new Map<string, { at: number; value: CrossrefWork | null }>();
const CACHE_TTL_MS = 60 * 60 * 1000;

function cached(key: string): CrossrefWork | null | undefined {
  const hit = cache.get(key);
  if (!hit) return undefined;
  if (Date.now() - hit.at > CACHE_TTL_MS) {
    cache.delete(key);
    return undefined;
  }
  return hit.value;
}

function normalizeTitle(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9 ]/g, "").replace(/\s+/g, " ").trim();
}

/** Simple token-overlap similarity, 0..1. */
function titleSimilarity(a: string, b: string): number {
  const ta = new Set(normalizeTitle(a).split(" ").filter((w) => w.length > 2));
  const tb = new Set(normalizeTitle(b).split(" ").filter((w) => w.length > 2));
  if (ta.size === 0 || tb.size === 0) return 0;
  let overlap = 0;
  for (const w of ta) if (tb.has(w)) overlap++;
  return overlap / Math.max(ta.size, tb.size);
}

function workToMetadata(work: CrossrefWork): VerifiedMetadata {
  return {
    doi: work.DOI,
    title: work.title?.[0],
    authors: work.author?.map((a) =>
      `${a.family ?? ""}${a.given ? `, ${a.given.split(/\s+/).map((g) => g.charAt(0) + ".").join(" ")}` : ""}`
    ),
    journal: work["container-title"]?.[0],
    year: work.issued?.["date-parts"]?.[0]?.[0]?.toString(),
    volume: work.volume,
    issue: work.issue,
    pages: work.page,
  };
}

export class CrossrefProvider implements MetadataProvider {
  readonly name = "crossref";
  private requestsThisRun = 0;

  resetBudget(): void {
    this.requestsThisRun = 0;
  }

  private async fetchJson(url: string): Promise<unknown | null> {
    if (this.requestsThisRun >= config.crossrefMaxRequestsPerRun) {
      throw new BudgetExceeded();
    }
    this.requestsThisRun++;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), config.crossrefTimeoutMs);
    try {
      const headers: Record<string, string> = {
        "User-Agent": `APA7DocumentAuditor/1.0${config.crossrefMailto ? ` (mailto:${config.crossrefMailto})` : ""}`,
      };
      const res = await fetch(url, { signal: controller.signal, headers });
      if (res.status === 404) return null;
      if (!res.ok) throw new Error(`Crossref HTTP ${res.status}`);
      return await res.json();
    } finally {
      clearTimeout(timer);
    }
  }

  private async byDoi(doi: string): Promise<CrossrefWork | null> {
    const key = `doi:${doi.toLowerCase()}`;
    const hit = cached(key);
    if (hit !== undefined) return hit;
    const json = (await this.fetchJson(
      `${config.crossrefBaseUrl}/works/${encodeURIComponent(doi)}`
    )) as { message?: CrossrefWork } | null;
    const work = json?.message ?? null;
    cache.set(key, { at: Date.now(), value: work });
    return work;
  }

  private async search(ref: ParsedReference): Promise<CrossrefWork[]> {
    const q = [ref.title, ref.authorsRaw, ref.year].filter(Boolean).join(" ");
    const key = `q:${q.toLowerCase()}`;
    const hit = cached(key);
    if (hit !== undefined) return hit ? [hit] : [];
    const json = (await this.fetchJson(
      `${config.crossrefBaseUrl}/works?query.bibliographic=${encodeURIComponent(q)}&rows=3&select=DOI,title,author,container-title,issued,volume,issue,page`
    )) as { message?: { items?: CrossrefWork[] } } | null;
    const items = json?.message?.items ?? [];
    cache.set(key, { at: Date.now(), value: items[0] ?? null });
    return items;
  }

  async verify(ref: ParsedReference): Promise<VerificationResult> {
    const base: Omit<VerificationResult, "status" | "confidence"> = {
      referenceIndex: ref.paragraphIndex,
      provider: this.name,
    };
    try {
      if (ref.doi) {
        const work = await this.byDoi(ref.doi);
        if (!work) {
          return {
            ...base,
            status: "mismatch",
            confidence: 0.9,
            note: `DOI ${ref.doi} was not found at Crossref. Check the DOI for typos.`,
          };
        }
        return this.compare(ref, work, 0.99);
      }
      if (!ref.title || ref.title.length < 8) {
        return {
          ...base,
          status: "unverified",
          confidence: 0,
          note: "Not enough bibliographic data (no DOI, short/missing title) to verify.",
        };
      }
      const candidates = await this.search(ref);
      let best: { work: CrossrefWork; sim: number } | null = null;
      for (const work of candidates) {
        const t = work.title?.[0];
        if (!t) continue;
        const sim = titleSimilarity(ref.title, t);
        if (!best || sim > best.sim) best = { work, sim };
      }
      if (!best || best.sim < 0.5) {
        return {
          ...base,
          status: "unverified",
          confidence: best?.sim ?? 0,
          note: "No sufficiently similar record found at Crossref. The source may not be indexed (books, webpages, reports often are not).",
        };
      }
      return this.compare(ref, best.work, Math.min(0.6 + best.sim * 0.4, 0.95));
    } catch (err) {
      if (err instanceof BudgetExceeded) {
        return {
          ...base,
          status: "unverified",
          confidence: 0,
          note: "Verification request budget for this run was reached.",
        };
      }
      return {
        ...base,
        status: "provider_unavailable",
        confidence: 0,
        note: "External metadata verification temporarily unavailable.",
      };
    }
  }

  private compare(
    ref: ParsedReference,
    work: CrossrefWork,
    baseConfidence: number
  ): VerificationResult {
    const metadata = workToMetadata(work);
    const differences: NonNullable<VerificationResult["differences"]> = [];

    if (ref.year && metadata.year && ref.year.replace(/[a-z]$/, "") !== metadata.year) {
      differences.push({
        field: "year",
        documentValue: ref.year,
        verifiedValue: metadata.year,
      });
    }
    if (ref.title && metadata.title) {
      const sim = titleSimilarity(ref.title, metadata.title);
      if (sim < 0.75) {
        differences.push({
          field: "title",
          documentValue: ref.title,
          verifiedValue: metadata.title,
        });
      }
    }
    if (ref.surnames.length > 0 && metadata.authors && metadata.authors.length > 0) {
      const refFirst = ref.surnames[0]!;
      const workFirst = (metadata.authors[0] ?? "").split(",")[0]!.toLowerCase();
      if (workFirst && refFirst !== workFirst) {
        differences.push({
          field: "first author",
          documentValue: refFirst,
          verifiedValue: workFirst,
        });
      }
    }
    if (ref.journal?.volume && metadata.volume && ref.journal.volume !== metadata.volume) {
      differences.push({
        field: "volume",
        documentValue: ref.journal.volume,
        verifiedValue: metadata.volume,
      });
    }

    if (differences.length === 0) {
      return {
        referenceIndex: ref.paragraphIndex,
        status: "verified",
        confidence: baseConfidence,
        provider: this.name,
        metadata,
      };
    }
    const severe = differences.some((d) => d.field === "title" || d.field === "first author");
    return {
      referenceIndex: ref.paragraphIndex,
      status: severe ? "mismatch" : "probable",
      confidence: severe ? 0.8 : baseConfidence * 0.85,
      provider: this.name,
      metadata,
      differences,
      note: severe
        ? "The matched record differs significantly from the document entry. Review both values — nothing was changed automatically."
        : "Minor differences found between the document entry and the verified record.",
    };
  }
}

class BudgetExceeded extends Error {}

import { NullProvider } from "./provider.js";

let providerInstance: MetadataProvider | null = null;

export function getProvider(): MetadataProvider {
  if (!providerInstance) {
    providerInstance =
      config.metadataProvider === "crossref"
        ? new CrossrefProvider()
        : new NullProvider();
  }
  return providerInstance;
}

/** For tests. */
export function setProvider(p: MetadataProvider | null): void {
  providerInstance = p;
}
