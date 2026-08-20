import type { DocxPackage } from "../docx/package.js";
import { NS, type XElement } from "../docx/xml.js";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/**
 * File extensions pdfkit can embed natively (pure JS, no native decoders).
 * Anything else a .docx might contain (GIF, BMP, TIFF, EMF/WMF vector
 * drawings, SVG) is skipped rather than crashing the merge — those formats
 * would need a native image decoder to rasterize, which is exactly the kind
 * of native-binary dependency this rewrite exists to avoid.
 */
const SUPPORTED_EXTENSION_BY_CONTENT: Record<string, true> = {
  png: true,
  jpg: true,
  jpeg: true,
};

const relationshipCache = new WeakMap<DocxPackage, Map<string, string>>();

async function relationshipTargets(pkg: DocxPackage): Promise<Map<string, string>> {
  const cached = relationshipCache.get(pkg);
  if (cached) return cached;
  const map = new Map<string, string>();
  const relsDoc = await pkg.getXmlIfPresent("word/_rels/document.xml.rels");
  if (relsDoc?.documentElement) {
    const rels = relsDoc.documentElement.getElementsByTagNameNS(NS.rels, "Relationship");
    for (let i = 0; i < rels.length; i++) {
      const rel = rels.item(i) as XElement;
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      if (id && target) map.set(id, target);
    }
  }
  relationshipCache.set(pkg, map);
  return map;
}

/**
 * Extracts every PNG/JPEG image embedded in a single run's w:drawing
 * element(s), resolved through the document's relationship part. Runs whose
 * only drawing is an unsupported format (or whose target is missing) yield
 * an empty array rather than throwing, so a single unusual image never
 * aborts the whole merge.
 */
export async function extractRunImages(pkg: DocxPackage, run: XElement): Promise<Buffer[]> {
  const drawings = run.getElementsByTagNameNS(NS.w, "drawing");
  if (drawings.length === 0) return [];
  const targets = await relationshipTargets(pkg);
  const out: Buffer[] = [];
  for (let i = 0; i < drawings.length; i++) {
    const drawing = drawings.item(i) as XElement;
    const blips = drawing.getElementsByTagNameNS(DRAWINGML_NS, "blip");
    for (let j = 0; j < blips.length; j++) {
      const blip = blips.item(j) as XElement;
      const rId = blip.getAttributeNS(NS.r, "embed");
      if (!rId) continue;
      const target = targets.get(rId);
      if (!target) continue;
      const ext = target.split(".").pop()?.toLowerCase() ?? "";
      if (!SUPPORTED_EXTENSION_BY_CONTENT[ext]) continue;
      const partName = target.startsWith("/") ? target.slice(1) : `word/${target}`;
      if (!pkg.has(partName)) continue;
      try {
        out.push(await pkg.readBinary(partName));
      } catch {
        /* unreadable media part — skip rather than fail the merge */
      }
    }
  }
  return out;
}
