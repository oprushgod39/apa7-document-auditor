import JSZip from "jszip";
import { Errors } from "../errors.js";
import { config } from "../config.js";
import { parseXml, serializeXml, type XDocument } from "./xml.js";

/**
 * Safe wrapper around a .docx OOXML package.
 *
 * Design principle: the ZIP is the source of truth. Parts we never touch are
 * carried through to the output byte-for-byte-equivalent; only parts we
 * explicitly modify are re-serialized. This is what preserves images,
 * equations, embedded media, custom XML, themes, fonts, and anything else
 * python-docx-style libraries tend to drop.
 */
export class DocxPackage {
  private zip: JSZip;
  private parsed = new Map<string, XDocument>();
  private dirty = new Set<string>();

  private constructor(zip: JSZip) {
    this.zip = zip;
  }

  static async load(buffer: Buffer): Promise<DocxPackage> {
    // OLE compound file signature => encrypted/password-protected Office file
    if (
      buffer.length >= 8 &&
      buffer[0] === 0xd0 &&
      buffer[1] === 0xcf &&
      buffer[2] === 0x11 &&
      buffer[3] === 0xe0
    ) {
      throw Errors.passwordProtected();
    }
    if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
      throw Errors.corrupt();
    }
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch {
      throw Errors.corrupt();
    }

    // --- Package safety validation -------------------------------------
    const names = Object.keys(zip.files);
    if (names.length > config.zip.maxEntries) {
      throw Errors.unsafePackage("too many entries in package");
    }
    let total = 0;
    for (const name of names) {
      if (name.includes("..") || name.startsWith("/") || /^[a-zA-Z]:/.test(name)) {
        throw Errors.unsafePackage("unsafe entry path");
      }
      const entry = zip.files[name]!;
      // jszip internal: _data.uncompressedSize when available
      const size =
        (entry as unknown as { _data?: { uncompressedSize?: number } })._data
          ?.uncompressedSize ?? 0;
      if (size > config.zip.maxEntryUncompressedBytes) {
        throw Errors.unsafePackage("entry exceeds decompression limit");
      }
      total += Math.max(0, size);
      if (total > config.zip.maxTotalUncompressedBytes) {
        throw Errors.unsafePackage("package exceeds decompression limit");
      }
      if (/vbaProject/i.test(name)) {
        throw Errors.macroRejected();
      }
    }

    if (!zip.file("[Content_Types].xml") || !zip.file("word/document.xml")) {
      throw Errors.corrupt();
    }

    const pkg = new DocxPackage(zip);
    const ct = await pkg.readText("[Content_Types].xml");
    if (/macroEnabled/i.test(ct)) {
      throw Errors.macroRejected();
    }
    // Validate the main document parses at all.
    await pkg.getXml("word/document.xml");
    return pkg;
  }

  has(partName: string): boolean {
    return this.zip.file(partName) != null;
  }

  partNames(): string[] {
    return Object.keys(this.zip.files).filter((n) => !this.zip.files[n]!.dir);
  }

  async readText(partName: string): Promise<string> {
    const f = this.zip.file(partName);
    if (!f) throw new Error(`Missing part: ${partName}`);
    return f.async("string");
  }

  async readBinary(partName: string): Promise<Buffer> {
    const f = this.zip.file(partName);
    if (!f) throw new Error(`Missing part: ${partName}`);
    return Buffer.from(await f.async("uint8array"));
  }

  /** Parse a part as XML (cached). Throws on malformed XML / DOCTYPE. */
  async getXml(partName: string): Promise<XDocument> {
    const cached = this.parsed.get(partName);
    if (cached) return cached;
    let text: string;
    try {
      text = await this.readText(partName);
    } catch {
      throw Errors.corrupt();
    }
    let doc: XDocument;
    try {
      doc = parseXml(text, partName);
    } catch {
      throw Errors.corrupt();
    }
    this.parsed.set(partName, doc);
    return doc;
  }

  /** Parse if present, else null. */
  async getXmlIfPresent(partName: string): Promise<XDocument | null> {
    if (!this.has(partName)) return null;
    return this.getXml(partName);
  }

  /** Mark a parsed part as modified so it is re-serialized on save. */
  markDirty(partName: string): void {
    if (!this.parsed.has(partName)) {
      throw new Error(`Cannot mark unparsed part dirty: ${partName}`);
    }
    this.dirty.add(partName);
  }

  /** Add or replace a part with raw text content (e.g. new header part). */
  setTextPart(partName: string, content: string): void {
    this.zip.file(partName, content);
    this.parsed.delete(partName);
    this.dirty.delete(partName);
  }

  get dirtyParts(): ReadonlySet<string> {
    return this.dirty;
  }

  async save(): Promise<Buffer> {
    for (const partName of this.dirty) {
      const doc = this.parsed.get(partName);
      if (doc) this.zip.file(partName, serializeXml(doc));
    }
    const out = await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    return Buffer.from(out);
  }
}

/**
 * Post-generation corruption check: reopen the produced buffer, ensure the
 * ZIP is valid, required parts exist and all XML parts parse. Never ship a
 * file that fails this.
 */
export async function verifyDocxIntegrity(buffer: Buffer): Promise<void> {
  const pkg = await DocxPackage.load(buffer);
  const required = ["[Content_Types].xml", "word/document.xml"];
  for (const part of required) {
    if (!pkg.has(part)) throw new Error(`Generated file missing ${part}`);
  }
  for (const name of pkg.partNames()) {
    if (name.endsWith(".xml") || name.endsWith(".rels")) {
      const text = await pkg.readText(name);
      if (text.trim().length === 0) continue;
      parseXml(text, name);
    }
  }
}
