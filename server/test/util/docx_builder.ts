import JSZip from "jszip";

/**
 * Test fixture builder: constructs genuine .docx packages (OOXML ZIPs) with
 * controllable formatting defects for golden-document tests.
 */

export interface ParaSpec {
  /** When set, used verbatim as the `<w:p>...</w:p>` XML — every other field is ignored. */
  rawXml?: string;
  text?: string;
  style?: string;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right" | "both";
  font?: string;
  sizeHalf?: number;
  spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  indent?: { firstLine?: number; hanging?: number; left?: number };
  pageBreakAfter?: boolean;
  leadingTab?: boolean;
  hyperlink?: { text: string; url: string };
  image?: boolean;
}

export interface DocSpec {
  paragraphs: ParaSpec[];
  margins?: { top: number; bottom: number; left: number; right: number };
  headerWithPageNumber?: boolean;
  table?: { rows: number; cols: number; cellText?: string };
  /** Insert the table after this paragraph index. */
  tableAfter?: number;
  defaultFont?: string;
  defaultSizeHalf?: number;
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function paraXml(p: ParaSpec): string {
  if (p.rawXml != null) return p.rawXml;
  const pPrParts: string[] = [];
  if (p.style) pPrParts.push(`<w:pStyle w:val="${p.style}"/>`);
  if (p.spacing) {
    const attrs = [
      p.spacing.before != null ? `w:before="${p.spacing.before}"` : "",
      p.spacing.after != null ? `w:after="${p.spacing.after}"` : "",
      p.spacing.line != null ? `w:line="${p.spacing.line}"` : "",
      p.spacing.lineRule != null ? `w:lineRule="${p.spacing.lineRule}"` : "",
    ]
      .filter(Boolean)
      .join(" ");
    pPrParts.push(`<w:spacing ${attrs}/>`);
  }
  if (p.indent) {
    const attrs = [
      p.indent.firstLine != null ? `w:firstLine="${p.indent.firstLine}"` : "",
      p.indent.hanging != null ? `w:hanging="${p.indent.hanging}"` : "",
      p.indent.left != null ? `w:left="${p.indent.left}"` : "",
    ]
      .filter(Boolean)
      .join(" ");
    pPrParts.push(`<w:ind ${attrs}/>`);
  }
  if (p.align) pPrParts.push(`<w:jc w:val="${p.align}"/>`);
  const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";

  const rPrParts: string[] = [];
  if (p.font) rPrParts.push(`<w:rFonts w:ascii="${p.font}" w:hAnsi="${p.font}"/>`);
  if (p.bold) rPrParts.push("<w:b/>");
  if (p.italic) rPrParts.push("<w:i/>");
  if (p.sizeHalf) rPrParts.push(`<w:sz w:val="${p.sizeHalf}"/>`);
  const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join("")}</w:rPr>` : "";

  const runs: string[] = [];
  if (p.leadingTab) runs.push(`<w:r>${rPr}<w:tab/></w:r>`);
  if (p.text != null && p.text.length > 0) {
    runs.push(`<w:r>${rPr}<w:t xml:space="preserve">${esc(p.text)}</w:t></w:r>`);
  }
  if (p.hyperlink) {
    runs.push(
      `<w:hyperlink r:id="rIdHl"><w:r>${rPr}<w:t>${esc(p.hyperlink.text)}</w:t></w:r></w:hyperlink>`
    );
  }
  if (p.image) {
    runs.push(
      `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    );
  }
  if (p.pageBreakAfter) runs.push(`<w:r><w:br w:type="page"/></w:r>`);
  return `<w:p>${pPr}${runs.join("")}</w:p>`;
}

function tableXml(rows: number, cols: number, cellText = "Cell"): string {
  let rowsXml = "";
  for (let r = 0; r < rows; r++) {
    let cells = "";
    for (let c = 0; c < cols; c++) {
      cells += `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${esc(cellText)} ${r},${c}</w:t></w:r></w:p></w:tc>`;
    }
    rowsXml += `<w:tr>${cells}</w:tr>`;
  }
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(cols)}</w:tblGrid>${rowsXml}</w:tbl>`;
}

// Minimal valid 1x1 transparent PNG
const PNG_BYTES = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
  "base64"
);

export async function buildDocx(spec: DocSpec): Promise<Buffer> {
  const zip = new JSZip();
  const m = spec.margins ?? { top: 1440, bottom: 1440, left: 1440, right: 1440 };

  const blocks: string[] = [];
  spec.paragraphs.forEach((p, i) => {
    blocks.push(paraXml(p));
    if (spec.table && spec.tableAfter === i) {
      blocks.push(tableXml(spec.table.rows, spec.table.cols, spec.table.cellText));
    }
  });
  if (spec.table && spec.tableAfter == null) {
    blocks.push(tableXml(spec.table.rows, spec.table.cols, spec.table.cellText));
  }

  const headerRef = spec.headerWithPageNumber
    ? `<w:headerReference w:type="default" r:id="rIdHdr"/>`
    : "";

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${blocks.join("")}<w:sectPr>${headerRef}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${spec.defaultFont ?? "Calibri"}" w:hAnsi="${spec.defaultFont ?? "Calibri"}"/><w:sz w:val="${spec.defaultSizeHalf ?? 22}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="160"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="4472C4" w:themeColor="accent1"/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="4472C4" w:themeColor="accent1"/><w:sz w:val="26"/></w:rPr></w:style>
</w:styles>`;

  const usesImage = spec.paragraphs.some((p) => p.image);
  const usesHyperlink = spec.paragraphs.some((p) => p.hyperlink);

  const relEntries: string[] = [
    `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
  ];
  if (spec.headerWithPageNumber) {
    relEntries.push(
      `<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`
    );
  }
  if (usesImage) {
    relEntries.push(
      `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`
    );
  }
  if (usesHyperlink) {
    const url = spec.paragraphs.find((p) => p.hyperlink)!.hyperlink!.url;
    relEntries.push(
      `<Relationship Id="rIdHl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`
    );
  }

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
${spec.headerWithPageNumber ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ""}
</Types>`;

  zip.file("[Content_Types].xml", contentTypes);
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntries.join("")}</Relationships>`
  );
  if (spec.headerWithPageNumber) {
    zip.file(
      "word/header1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`
    );
  }
  if (usesImage) zip.file("word/media/image1.png", PNG_BYTES);

  return Buffer.from(
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
  );
}

/** A typical malformed student paper used across golden tests. */
export function malformedStudentPaper(): DocSpec {
  return {
    defaultFont: "Calibri",
    defaultSizeHalf: 22,
    margins: { top: 1080, bottom: 1080, left: 1080, right: 1080 }, // 0.75"
    headerWithPageNumber: false,
    paragraphs: [
      // Title page
      { text: "The Effects of Sleep on Memory Consolidation", align: "center", bold: false },
      { text: "Jordan A. Rivera", align: "center" },
      { text: "Department of Psychology, Riverside University", align: "center" },
      { text: "PSY 301: Cognitive Psychology", align: "center" },
      { text: "Dr. Maya Chen", align: "center" },
      { text: "March 14, 2025", align: "center", pageBreakAfter: true },
      // Body
      {
        text: "Sleep is essential for cognitive performance. Smith and Patel (2024) found that sleep deprivation impairs recall. Other work has confirmed this pattern (Kumar et al., 2023). One study reported that participants remembered 40% fewer words (Smith & Patel, 2024, p. 21). A related result appears in earlier work (Johnson, 2023).",
        spacing: { after: 160, line: 259, lineRule: "auto" },
      },
      {
        text: "Method",
        bold: true,
        align: "center",
        spacing: { line: 259, lineRule: "auto" },
      },
      {
        text: "Participants completed a memory task after either a full night of sleep or total sleep deprivation. The design followed established protocols and all participants provided informed consent before beginning the experimental procedure in the laboratory.",
        leadingTab: true,
        spacing: { after: 160, line: 240, lineRule: "auto" },
        align: "both",
      },
      {
        text: "Results indicated a large effect of sleep on recall performance across both immediate and delayed testing conditions, and the pattern held for recognition memory as well.",
        spacing: { after: 160, line: 259, lineRule: "auto" },
        pageBreakAfter: true,
      },
      // References
      { text: "Bibliography", align: "left", bold: false },
      {
        text: "Smith, J. K., & Patel, R. (2024). Sleep deprivation and memory consolidation. Journal of Cognitive Neuroscience, 36(2), 112–128. doi:10.1234/jcn.2024.0112",
        indent: { firstLine: 720 },
      },
      {
        text: "Kumar, A., Lopez, M., & Chen, W. (2023). Sleep architecture and learning. Sleep Research, 12(4), 45–67. https://doi.org/10.5678/sr.2023.045",
        indent: {},
      },
      {
        text: "Adams, B. C. (2021). Foundations of sleep science (2nd ed.). Academic Press.",
        indent: {},
      },
    ],
  };
}
