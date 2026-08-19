import type { ParsedReference } from "../apa/references/parser.js";

/**
 * Metadata provider abstraction. Additional providers (DataCite, OpenAlex,
 * PubMed…) can be added by implementing this interface and registering it in
 * getProvider().
 */

export interface VerifiedMetadata {
  doi?: string;
  title?: string;
  authors?: string[]; // "Surname, I." strings
  journal?: string;
  year?: string;
  volume?: string;
  issue?: string;
  pages?: string;
}

export type VerificationStatus =
  | "verified"
  | "probable"
  | "mismatch"
  | "unverified"
  | "provider_unavailable";

export interface VerificationResult {
  referenceIndex: number;
  status: VerificationStatus;
  confidence: number; // 0..1
  provider: string;
  metadata?: VerifiedMetadata;
  /** Field-level differences: document value vs verified value. */
  differences?: { field: string; documentValue: string; verifiedValue: string }[];
  note?: string;
}

export interface MetadataProvider {
  readonly name: string;
  /** Verify a single parsed reference. Must not throw on provider failure. */
  verify(ref: ParsedReference): Promise<VerificationResult>;
}

export class NullProvider implements MetadataProvider {
  readonly name = "none";
  async verify(ref: ParsedReference): Promise<VerificationResult> {
    return {
      referenceIndex: ref.paragraphIndex,
      status: "unverified",
      confidence: 0,
      provider: this.name,
      note: "External metadata verification is disabled.",
    };
  }
}
