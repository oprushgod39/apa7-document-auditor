import { describe, expect, it } from "vitest";
import { webcrypto } from "node:crypto";
import {
  decryptMergePayload,
  validateMergeBlobUrl,
} from "../src/merge/blob_transport.js";

describe("merger blob transport", () => {
  const batchId = "79f79ec6-bbdd-4ec1-b814-2f9fb8aa0e8a";

  it("accepts only a Blob URL inside the requested merge batch", () => {
    const url = `https://store.public.blob.vercel-storage.com/merge-inputs/${batchId}/1-paper-random.docx`;
    expect(validateMergeBlobUrl(url, batchId, "merge-inputs")).toBe(url);
  });

  it("rejects another batch and lookalike hosts", () => {
    const other =
      "https://store.public.blob.vercel-storage.com/merge-inputs/8cd92217-6459-4577-ac3f-75d4b02a11f0/paper.docx";
    const lookalike = `https://blob.vercel-storage.com.attacker.example/merge-inputs/${batchId}/paper.docx`;
    expect(() => validateMergeBlobUrl(other, batchId, "merge-inputs")).toThrow();
    expect(() => validateMergeBlobUrl(lookalike, batchId, "merge-inputs")).toThrow();
  });

  it("decrypts browser-compatible AES-GCM payloads and rejects tampering", async () => {
    const key = await webcrypto.subtle.generateKey(
      { name: "AES-GCM", length: 256 },
      true,
      ["encrypt"]
    );
    const rawKey = Buffer.from(await webcrypto.subtle.exportKey("raw", key));
    const iv = webcrypto.getRandomValues(new Uint8Array(12));
    const plain = Buffer.from("Private student document contents");
    const encrypted = Buffer.from(
      await webcrypto.subtle.encrypt({ name: "AES-GCM", iv }, key, plain)
    );

    expect(
      decryptMergePayload(
        encrypted,
        rawKey.toString("base64url"),
        Buffer.from(iv).toString("base64url"),
        plain.length
      )
    ).toEqual(plain);

    encrypted[0] ^= 1;
    expect(() =>
      decryptMergePayload(
        encrypted,
        rawKey.toString("base64url"),
        Buffer.from(iv).toString("base64url"),
        plain.length
      )
    ).toThrow(/could not be decrypted/i);
  });
});
