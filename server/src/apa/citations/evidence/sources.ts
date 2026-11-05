import https from "node:https";
import { lookup } from "node:dns/promises";
import type { ParsedReference } from "../../references/parser.js";
import type { EvidenceSource } from "./types.js";

export function publicIPv4(ip: string): boolean {
  const p = ip.split(".").map(Number);
  if (p.length !== 4 || p.some((n) => !Number.isInteger(n) || n < 0 || n > 255)) return false;
  const [a, b] = p as [number, number, number, number];
  return !(a === 0 || a === 10 || a === 127 || a >= 224 ||
    (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && (b === 168 || b === 0 || b === 2)) || (a === 100 && b >= 64 && b <= 127) ||
    (a === 198 && (b === 18 || b === 19 || b === 51)) || (a === 203 && b === 0));
}

/** Pin DNS at connect time; revalidate every redirect. No credentials/cookies. */
export async function sourceGet(address: string, redirects = 0): Promise<{ url: string; text: string; data: Buffer; type: string }> {
  const url = new URL(address);
  if (url.protocol !== "https:" || url.username || url.password || (url.port && url.port !== "443") || redirects > 4)
    throw new Error("Unsupported source address.");
  const addresses = (await lookup(url.hostname, { family: 4, all: true })).map((a) => a.address);
  if (!addresses.length || addresses.some((ip) => !publicIPv4(ip))) throw new Error("Source address is not public.");
  return new Promise((resolve, reject) => {
    const req = https.get(url, { agent: false, family: 4, lookup: (_host, _options, callback) => callback(null, addresses[0]!, 4),
      headers: { "User-Agent": "ScholarlyWorks-CitationVerifier/1.0", Accept: "text/html,application/json,text/plain" } }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode ?? 0) && res.headers.location) {
        res.resume(); clearTimeout(timer);
        sourceGet(new URL(res.headers.location, url).href, redirects + 1).then(resolve, reject); return;
      }
      if (res.statusCode !== 200) { clearTimeout(timer); res.resume(); reject(new Error(`Source returned HTTP ${res.statusCode}.`)); return; }
      let bytes = 0; const chunks: Buffer[] = [];
      res.on("data", (chunk: Buffer) => { bytes += chunk.length; if (bytes > 2_000_000) req.destroy(new Error("Source exceeds retrieval limit.")); else chunks.push(chunk); });
      res.on("end", () => { clearTimeout(timer); const data = Buffer.concat(chunks); resolve({ url: url.href, text: data.toString("utf8"), data, type: String(res.headers["content-type"] ?? "") }); });
      res.on("error", reject);
    });
    const timer = setTimeout(() => req.destroy(new Error("Source retrieval timed out.")), 7000);
    req.on("error", (e) => { clearTimeout(timer); reject(e); });
  });
}

export function readable(text: string): string {
  return text.replace(/<(script|style|nav|footer|header)\b[^>]*>[\s\S]*?<\/\1>/gi, " ")
    .replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/\s+/g, " ").trim();
}

export async function retrieveSource(ref: ParsedReference, id: number): Promise<EvidenceSource> {
  const result: EvidenceSource = { id, reference: ref, url: ref.url ?? (ref.doi ? `https://doi.org/${ref.doi}` : ""), text: "", status: "unavailable", scope: "No accessible source content." };
  try {
    if (ref.doi) {
      const response = await sourceGet(`https://api.crossref.org/works/${encodeURIComponent(ref.doi)}`);
      const work = JSON.parse(response.text).message;
      const years = [work.published, work["published-print"], work["published-online"], work.issued]
        .flatMap((v) => v?.["date-parts"]?.[0]?.[0] ? [String(v["date-parts"][0][0])] : []);
      if (ref.year && years.length && !years.includes(ref.year.replace(/[a-z]$/, ""))) {
        result.status = "metadata_conflict"; result.scope = `Reference year ${ref.year} differs from source metadata (${years.join(", ")}).`; return result;
      }
      if (work.abstract) {
        result.text = `Title: ${work.title?.[0] ?? ""}. Authors: ${(work.author ?? []).map((a: { family?: string }) => a.family).join(", ")}. Abstract: ${readable(work.abstract)}`;
        result.url = response.url; result.status = "available"; result.scope = "Publisher-deposited abstract only; full text not examined.";
        return result;
      }
    }
    if (!result.url) return result;
    const page = await sourceGet(result.url);
    if (/application\/pdf/i.test(page.type)) {
      const pdfjs = await import("pdfjs-dist/legacy/build/pdf.js");
      const loading = pdfjs.getDocument({ data: new Uint8Array(page.data), isEvalSupported: false, disableFontFace: true, useSystemFonts: false });
      const pdf = await loading.promise;
      try {
        let text = ""; let pages = 0;
        for (let p = 1; p <= pdf.numPages && p <= 20 && text.length < 16000; p++) {
          const pageText = await (await pdf.getPage(p)).getTextContent();
          text += ` [PDF page ${p}] ` + pageText.items.map((item) => "str" in item ? item.str : "").join(" "); pages++;
        }
        result.text = text.slice(0, 16000); result.status = text.trim().length > 200 ? "available" : "unavailable";
        result.url = page.url; result.scope = `PDF extraction: ${pages} of ${pdf.numPages} pages, up to 16,000 characters. PDF page indexes are not verified printed page numbers. Scanned pages need manual review.`;
        return result;
      } finally { await loading.destroy(); }
    }
    if (!/text\/(?:html|plain)|application\/xhtml/i.test(page.type)) { result.scope = "Source format cannot be read automatically."; return result; }
    const text = readable(page.text);
    if (text.length < 200 || /^(?:access denied|just a moment|verify you are human)/i.test(text)) return result;
    result.text = text.slice(0, 16000); result.status = "available"; result.url = page.url;
    result.scope = text.length > 16000 ? "First 16,000 characters of source page; remaining content not examined." : "Accessible source webpage; identity must be checked against the reference.";
  } catch { result.scope = "Source retrieval failed or access was restricted. No verification claimed."; }
  return result;
}
