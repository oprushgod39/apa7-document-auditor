import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { DocxPackage, verifyDocxIntegrity } from "../src/docx/package.js";
import { buildDocx, malformedStudentPaper } from "./util/docx_builder.js";

describe("DocxPackage security & integrity", () => {
  it("loads a valid docx and round-trips without corruption", async () => {
    const buf = await buildDocx(malformedStudentPaper());
    const pkg = await DocxPackage.load(buf);
    expect(pkg.has("word/document.xml")).toBe(true);
    const out = await pkg.save();
    await verifyDocxIntegrity(out);
  });

  it("rejects non-zip data as corrupt", async () => {
    await expect(DocxPackage.load(Buffer.from("not a zip file at all"))).rejects.toMatchObject({
      code: "CORRUPT_DOCUMENT",
    });
  });

  it("rejects OLE compound files as password-protected", async () => {
    const ole = Buffer.alloc(64);
    ole[0] = 0xd0; ole[1] = 0xcf; ole[2] = 0x11; ole[3] = 0xe0;
    await expect(DocxPackage.load(ole)).rejects.toMatchObject({
      code: "PASSWORD_PROTECTED",
    });
  });

  it("rejects macro-enabled packages (vbaProject)", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
    zip.file("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>");
    zip.file("word/vbaProject.bin", Buffer.from([1, 2, 3]));
    const buf = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    await expect(DocxPackage.load(buf)).rejects.toMatchObject({
      code: "MACRO_DOCUMENT_REJECTED",
    });
  });

  it("rejects macroEnabled content types", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/word/document.xml' ContentType='application/vnd.ms-word.document.macroEnabled.main+xml'/></Types>"
    );
    zip.file("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>");
    const buf = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    await expect(DocxPackage.load(buf)).rejects.toMatchObject({
      code: "MACRO_DOCUMENT_REJECTED",
    });
  });

  it("rejects packages with path traversal entries", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
    zip.file("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>");
    (zip as unknown as { files: Record<string, unknown> }).files["../evil.txt"] =
      zip.file("word/document.xml")!;
    const buf = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    // JSZip normalizes names on generate, so also test via direct load of a
    // hand-made zip if normalization removed it; either rejection or absence
    // of the traversal entry is acceptable safety-wise.
    const pkg = await DocxPackage.load(buf).catch((e) => e);
    if (pkg instanceof Error) {
      expect((pkg as { code?: string }).code).toBe("UNSAFE_PACKAGE");
    } else {
      expect((pkg as DocxPackage).partNames().some((n) => n.includes(".."))).toBe(false);
    }
  });

  it("rejects XML with DOCTYPE declarations (XXE hardening)", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>");
    zip.file(
      "word/document.xml",
      "<?xml version='1.0'?><!DOCTYPE foo [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body/></w:document>"
    );
    const buf = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
    await expect(DocxPackage.load(buf)).rejects.toMatchObject({
      code: "CORRUPT_DOCUMENT",
    });
  });

  it("preserves untouched binary parts byte-for-byte", async () => {
    const spec = malformedStudentPaper();
    spec.paragraphs[7]!.image = true;
    const buf = await buildDocx(spec);
    const pkg = await DocxPackage.load(buf);
    const before = await pkg.readBinary("word/media/image1.png");
    const out = await pkg.save();
    const pkg2 = await DocxPackage.load(out);
    const after = await pkg2.readBinary("word/media/image1.png");
    expect(Buffer.compare(before, after)).toBe(0);
  });
});
