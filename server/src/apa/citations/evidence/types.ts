import type { ParsedReference } from "../../references/parser.js";

export interface EvidenceSource {
  id: number;
  reference: ParsedReference;
  url: string;
  text: string;
  status: "available" | "unavailable" | "metadata_conflict";
  scope: string;
}
export interface ClaimDecision {
  sentence: string;
  claim: string;
  category: "research" | "statistic" | "quotation" | "common_knowledge" | "original_analysis" | "other";
  status: "supported" | "partial" | "contradicted" | "insufficient" | "no_citation_needed";
  reason: string;
  evidence: { sourceId: number; quote: string; identityConfirmed: boolean }[];
  action: "none" | "insert" | "replace" | "remove";
  originalCitation: string;
  anchor: string;
  sourceIds: number[];
}
export interface ClaimReport extends ClaimDecision {
  paragraph: number;
  correctedCitation: string;
  applied: boolean;
  review: string;
}
export interface EvidenceJob {
  tokenHash: string;
  version: number;
  phase: "sources" | "claims" | "finalize" | "complete";
  sourceCursor: number;
  paragraphCursor: number;
  references: ParsedReference[];
  sources: EvidenceSource[];
  paragraphs: { index: number; text: string; editable: boolean }[];
  claims: ClaimReport[];
  edits: { paragraph: number; start: number; length: number; value: string }[];
  mechanics: { paragraph: number; before: string; after: string; reason: string }[];
  warnings: string[];
  uncitedReferences: string[];
  model: string;
  pending?: ClaimDecision[];
}
