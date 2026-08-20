export interface SimilarityDocument {
  name: string;
  display: string;
  analysis: string;
}

export interface SimilarityResult {
  i: number;
  j: number;
  a: string;
  b: string;
  phrase: number;
  content: number;
  overall: number;
}

export interface PassageMatch {
  a: string;
  b: string;
  score: number;
  shared: number;
}

const SETTINGS = {
  phraseWords: 8,
  passagePhraseWords: 6,
  passageThreshold: 0.58,
  contentWeight: 0.7,
  phraseWeight: 0.3,
  boilerplateFraction: 0.8,
  maxPassagesPerDoc: 900,
  maxDisplayedMatches: 25,
} as const;

const STOP = new Set(
  "the a an and or but if then than to of in on for with as at by from is are was were be been being it this that these those we you they he she i our your their his her its not can could should would may might will do does did have has had into about over under between such using use used also more most some any all each other another which who what when where why how through within without during per via among because therefore however thus while whereupon upon there here very much many few both either neither same own just only".split(" ")
);

function words(text: string): string[] {
  return text.toLocaleLowerCase().match(/[\p{L}\p{N}']+/gu) ?? [];
}

function stripReferences(text: string): string {
  const heading = /(?:^|\n)\s*(references|reference list|bibliography|works cited)\s*(?:\n|$)/gi;
  const candidates: number[] = [];
  let match: RegExpExecArray | null;
  while ((match = heading.exec(text)) != null) candidates.push(match.index);
  const later = candidates.filter((index) => index >= text.length * 0.3);
  if (later.length) return text.slice(0, later[0]);

  const tailStart = Math.floor(text.length * 0.55);
  const tail = text.slice(tailStart);
  const flattened = /\breferences\b/gi;
  while ((match = flattened.exec(tail)) != null) {
    const absolute = tailStart + match.index;
    const after = text.slice(absolute + match[0].length, absolute + match[0].length + 900);
    const evidence = (after.match(/\(?(?:19|20)\d{2}[a-z]?\)?/g) ?? []).length +
      (after.match(/doi\b|https?:\/\//gi) ?? []).length;
    if (evidence >= 2) return text.slice(0, absolute);
  }
  return text;
}

function prepareDisplay(raw: string): string {
  return stripReferences(
    raw.normalize("NFKC").replace(/\u00ad/g, "").replace(/\r\n?/g, "\n")
      .replace(/[\t ]+/g, " ").replace(/ *\n */g, "\n").replace(/\n{3,}/g, "\n\n")
  ).trim();
}

function normalizeForAnalysis(text: string): string {
  return text.toLocaleLowerCase().replace(/\[page\s+\d+]/gi, " ")
    .replace(/[’‘`]/g, "'").replace(/[“”]/g, '"')
    .replace(/[^\p{L}\p{N}'\s]+/gu, " ").replace(/\s+/g, " ").trim();
}

async function extractFile(file: File): Promise<string> {
  const extension = file.name.toLocaleLowerCase().split(".").pop();
  if (extension === "txt") return file.text();
  if (extension === "docx") {
    const mammoth = (await import("mammoth")).default;
    const output = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
    return output.value ?? "";
  }
  if (extension === "pdf") {
    const pdfjsLib = await import("pdfjs-dist");
    const pdfWorker = (await import("pdfjs-dist/build/pdf.worker.min.mjs?url")).default;
    pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let text = "";
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      text += `\n\n[PAGE ${pageNumber}]\n` + content.items
        .map((item) => ("str" in item ? item.str : "")).join(" ");
    }
    return text;
  }
  throw new Error(`Unsupported file: ${file.name}`);
}

function contentWords(text: string): string[] {
  return words(text).filter((word) => word.length > 2 && !STOP.has(word));
}

function ngramSet(text: string, size: number): Set<string> {
  const tokens = words(text);
  const set = new Set<string>();
  for (let index = 0; index <= tokens.length - size; index++) {
    set.add(tokens.slice(index, index + size).join(" "));
  }
  return set;
}

function featureFrequency(text: string): Map<string, number> {
  const tokens = contentWords(text);
  const map = new Map<string, number>();
  for (const token of tokens) map.set(`u:${token}`, (map.get(`u:${token}`) ?? 0) + 1);
  for (let index = 0; index < tokens.length - 1; index++) {
    const feature = `b:${tokens[index]} ${tokens[index + 1]}`;
    map.set(feature, (map.get(feature) ?? 0) + 1);
  }
  return map;
}

function buildTfidf(texts: string[]): Map<string, number>[] {
  const termFrequencies = texts.map(featureFrequency);
  const documentFrequency = new Map<string, number>();
  for (const frequency of termFrequencies) {
    for (const key of frequency.keys()) documentFrequency.set(key, (documentFrequency.get(key) ?? 0) + 1);
  }
  const ubiquitousCutoff = texts.length >= 5 ? Math.ceil(texts.length * 0.9) : Number.POSITIVE_INFINITY;
  return termFrequencies.map((frequency) => {
    const vector = new Map<string, number>();
    for (const [key, count] of frequency) {
      const seen = documentFrequency.get(key) ?? 1;
      if (seen >= ubiquitousCutoff) continue;
      vector.set(key, (1 + Math.log(count)) * (Math.log((texts.length + 1) / (seen + 1)) + 1));
    }
    return vector;
  });
}

function cosine(a: Map<string, number>, b: Map<string, number>): number {
  let dot = 0, magnitudeA = 0, magnitudeB = 0;
  for (const [key, value] of a) { magnitudeA += value * value; if (b.has(key)) dot += value * b.get(key)!; }
  for (const value of b.values()) magnitudeB += value * value;
  return magnitudeA && magnitudeB ? dot / Math.sqrt(magnitudeA * magnitudeB) : 0;
}

export async function compareFiles(
  files: File[],
  progress: (message: string) => void
): Promise<{ documents: SimilarityDocument[]; results: SimilarityResult[] }> {
  const documents: SimilarityDocument[] = [];
  for (let index = 0; index < files.length; index++) {
    progress(`Reading ${index + 1}/${files.length}: ${files[index]!.name}`);
    const display = prepareDisplay(await extractFile(files[index]!));
    const analysis = normalizeForAnalysis(display);
    if (words(analysis).length < 25) throw new Error(`${files[index]!.name} contains too little extractable text.`);
    documents.push({ name: files[index]!.name, display, analysis });
  }

  const texts = documents.map((document) => document.analysis);
  const phraseSets = texts.map((text) => ngramSet(text, SETTINGS.phraseWords));
  const phraseFrequency = new Map<string, number>();
  for (const set of phraseSets) for (const phrase of set) phraseFrequency.set(phrase, (phraseFrequency.get(phrase) ?? 0) + 1);
  const boilerplate = new Set<string>();
  if (texts.length >= 5) {
    const cutoff = Math.max(4, Math.ceil(texts.length * SETTINGS.boilerplateFraction));
    for (const [phrase, frequency] of phraseFrequency) if (frequency >= cutoff) boilerplate.add(phrase);
  }
  const vectors = buildTfidf(texts);
  const total = documents.length * (documents.length - 1) / 2;
  let completed = 0;
  const results: SimilarityResult[] = [];
  for (let i = 0; i < documents.length; i++) {
    for (let j = i + 1; j < documents.length; j++) {
      progress(`Comparing ${++completed}/${total}: ${documents[i]!.name} ↔ ${documents[j]!.name}`);
      const a = phraseSets[i]!, b = phraseSets[j]!;
      let countA = 0, countB = 0, intersection = 0;
      for (const phrase of a) if (!boilerplate.has(phrase)) { countA++; if (b.has(phrase)) intersection++; }
      for (const phrase of b) if (!boilerplate.has(phrase)) countB++;
      const phrase = countA && countB ? 2 * intersection / (countA + countB) : 0;
      const content = cosine(vectors[i]!, vectors[j]!);
      const overall = Math.min(1, SETTINGS.contentWeight * content + SETTINGS.phraseWeight * phrase);
      results.push({ i, j, a: documents[i]!.name, b: documents[j]!.name, phrase, content, overall });
      if (completed % 8 === 0) await new Promise((resolve) => window.setTimeout(resolve, 0));
    }
  }
  results.sort((left, right) => right.overall - left.overall);
  return { documents, results };
}

function splitSentences(text: string): string[] {
  return text.split(/(?<=[.!?])\s+(?=[A-Z0-9\[])/).map((sentence) => sentence.trim()).filter(Boolean);
}

function passages(text: string): string[] {
  const cleaned = text.replace(/\[PAGE\s+\d+]/gi, " ").trim();
  let output: string[] = [];
  const addChunk = (paragraph: string) => {
    if (paragraph.length <= 850) { output.push(paragraph); return; }
    let current = "";
    for (const sentence of splitSentences(paragraph)) {
      if (current && current.length + sentence.length > 650) { output.push(current.trim()); current = sentence; }
      else current += `${current ? " " : ""}${sentence}`;
    }
    if (current.trim().length >= 60) output.push(current.trim());
  };
  cleaned.split(/\n{2,}/).map((part) => part.trim()).filter((part) => part.length >= 60).forEach(addChunk);
  if (output.length < 4) {
    output = [];
    let current = "";
    for (const sentence of splitSentences(cleaned)) {
      if (current && current.length + sentence.length > 600) { if (current.length >= 60) output.push(current.trim()); current = sentence; }
      else current += `${current ? " " : ""}${sentence}`;
    }
    if (current.length >= 60) output.push(current.trim());
  }
  return output.slice(0, SETTINGS.maxPassagesPerDoc);
}

function tokenJaccard(a: string, b: string): number {
  const left = new Set(contentWords(normalizeForAnalysis(a)));
  const right = new Set(contentWords(normalizeForAnalysis(b)));
  if (!left.size || !right.size) return 0;
  let intersection = 0;
  for (const token of left) if (right.has(token)) intersection++;
  return intersection / (left.size + right.size - intersection);
}

function sharedPhraseCount(a: string, b: string): number {
  const left = ngramSet(normalizeForAnalysis(a), SETTINGS.passagePhraseWords);
  const right = ngramSet(normalizeForAnalysis(b), SETTINGS.passagePhraseWords);
  let shared = 0;
  for (const phrase of left) if (right.has(phrase)) shared++;
  return shared;
}

export async function comparePassages(a: SimilarityDocument, b: SimilarityDocument): Promise<PassageMatch[]> {
  const left = passages(a.display), right = passages(b.display), candidates: Array<PassageMatch & { rightIndex: number }> = [];
  for (let i = 0; i < left.length; i++) {
    let best: (PassageMatch & { rightIndex: number }) | null = null;
    for (let j = 0; j < right.length; j++) {
      const frequencyA = featureFrequency(normalizeForAnalysis(left[i]!));
      const frequencyB = featureFrequency(normalizeForAnalysis(right[j]!));
      const shared = sharedPhraseCount(left[i]!, right[j]!);
      let score = 0.75 * cosine(frequencyA, frequencyB) + 0.25 * tokenJaccard(left[i]!, right[j]!);
      if (shared > 0) score = Math.max(score, Math.min(1, 0.62 + 0.055 * shared));
      if (!best || score > best.score) best = { a: left[i]!, b: right[j]!, score, shared, rightIndex: j };
    }
    if (best && best.score >= SETTINGS.passageThreshold) candidates.push(best);
    if (i % 60 === 0) await new Promise((resolve) => window.setTimeout(resolve, 0));
  }
  candidates.sort((x, y) => y.score - x.score);
  const used = new Set<number>();
  return candidates.filter((candidate) => {
    if (used.has(candidate.rightIndex)) return false;
    used.add(candidate.rightIndex);
    return true;
  }).slice(0, SETTINGS.maxDisplayedMatches);
}

export function band(score: number): { label: string; className: string } {
  if (score >= 0.7) return { label: "Very high", className: "very-high" };
  if (score >= 0.5) return { label: "High", className: "high" };
  if (score >= 0.3) return { label: "Moderate", className: "moderate" };
  if (score >= 0.15) return { label: "Low", className: "low" };
  return { label: "Very low", className: "very-low" };
}

export const percent = (value: number) => `${(100 * value).toFixed(1)}%`;
