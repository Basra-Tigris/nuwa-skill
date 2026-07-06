// build_bilingual_docx.js
// docx-js template for generating bilingual Word documents:
//   - EN-AR file: left column English, right column Arabic
//   - ZH-AR file: left column Chinese, right column Arabic
//
// All body paragraphs use 1.2x line spacing (see LINE_SPACING_1_2 below).
//
// Usage:
//   1. Copy this file to your working directory.
//   2. Edit the DOCUMENTS array below with your content.
//   3. Install docx locally if needed (NEVER globally):
//        pnpm add docx
//   4. Run with NODE_PATH set when using a bundled package directory:
//        $env:NODE_PATH = "<NODE_MODULES>"
//        node build_bilingual_docx.js
//
// Output: one .docx per entry in DOCUMENTS, written to the current directory.

const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
  Header,
  Footer,
  PageNumber,
  ShadingType,
  BorderStyle,
  LevelFormat,
  LineRuleType,
  convertInchesToTwip,
} = require("docx");

// 1.2x line spacing for every paragraph in the document.
// When lineRule = AUTO, `line` is measured in 240ths of a single line,
// so 288 = 240 * 1.2 = 1.2x line spacing.
const LINE_SPACING_1_2 = { line: 288, lineRule: LineRuleType.AUTO };

// =============================================================================
// CONTENT — EDIT THIS SECTION
// =============================================================================
// Each document is one bilingual file. Pair each source/translation segment
// row-by-row in the `rows` array: [leftText, rightText].
// `mode` controls which pair: "EN-AR" (English left / Arabic right)
//                            "ZH-AR" (Chinese left / Arabic right)

const DOCUMENTS = [
  {
    mode: "EN-AR",
    basename: "sample_EN-AR",
    title: "Bilingual Document — English / Arabic",
    subtitle: "Working translation — not certified",
    confidentiality: "DRAFT — CONFIDENTIAL",
    rows: [
      ["Decision No. DE-1447-00053", "قرار رقم DE-1447-00053"],
      ["The General Authority for Competition, after reviewing the file,", "اطلعت الهيئة العامة للمنافسة على الملف،"],
      ["decides the following:", "تقرر ما يلي:"],
      ["First: The respondent shall cease the conduct.", "أولاً: يلتزم المخالف بإيقاف السلوك."],
      ["Second: A fine of SAR 100,000 is imposed.", "ثانياً: تُفرض غرامة قدرها 100,000 ريال سعودي."],
    ],
    translatorsNote: [
      "Translator's note (for counsel verification — this is a working translation, not a certified/sworn translation):",
      "1. Judgment calls: none in this sample.",
      "2. Ambiguities / source defects: none observed.",
      "3. Date conversions: none in this sample.",
      "4. Terminology choices: 'General Authority for Competition' used throughout for الهيئة العامة للمنافسة (GAC).",
      "5. Backend used: docx-js.",
    ],
  },
  {
    mode: "ZH-AR",
    basename: "sample_ZH-AR",
    title: "对照文档 — 中文 / 阿拉伯语",
    subtitle: "工作译本 — 非公证翻译",
    confidentiality: "草案 — 机密",
    rows: [
      ["第 DE-1447-00053 号决定", "قرار رقم DE-1447-00053"],
      ["竞争总局在审阅案卷后，", "اطلعت الهيئة العامة للمنافسة على الملف،"],
      ["决定如下：", "تقرر ما يلي:"],
      ["一、被调查方应停止上述行为。", "أولاً: يلتزم المخالف بإيقاف السلوك."],
      ["二、处以 100,000 沙特里亚尔罚款。", "ثانياً: تُفرض غرامة قدرها 100,000 ريال سعودي."],
    ],
    translatorsNote: [
      "译注（供律师复核 —— 本译本为工作翻译，非公证/宣誓翻译）：",
      "1. 判断取舍：本样例无。",
      "2. 歧义/源文瑕疵：未发现。",
      "3. 日期换算：本样例无。",
      "4. 术语选择：الهيئة العامة للمنافسة 全文统一译为\"竞争总局\"（GAC）。",
      "5. 使用的后端：docx-js。",
    ],
  },
];

// =============================================================================
// FONTS AND SIZES
// =============================================================================

const FONT_LATIN = "Times New Roman"; // English
const FONT_CJK = "SimSun"; // Chinese (宋体)
const FONT_ARABIC = "Amiri"; // Arabic — install Amiri font on the system, or fall back to "Traditional Arabic" / "Simplified Arabic"
const SIZE_BODY = 24; // half-points; 24 = 12pt
const SIZE_HEADING = 32; // 16pt
const SIZE_SUBTITLE = 22; // 11pt
const SIZE_NOTE = 20; // 10pt

// =============================================================================
// RUN / PARAGRAPH BUILDERS
// =============================================================================

function latinRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT_LATIN,
    size: opts.size || SIZE_BODY,
    bold: opts.bold || false,
    italics: opts.italics || false,
  });
}

function cjkRun(text, opts = {}) {
  // IMPORTANT: docx-js's RunFonts fills ALL font slots (ascii/cs/eastAsia/hAnsi)
  // from the `name` key and IGNORES a separately-supplied `eastAsia` key. So if
  // you write `font: { name: FONT_LATIN, eastAsia: FONT_CJK }`, every CJK run
  // ends up with eastAsia="Times New Roman" and Chinese renders via font
  // fallback instead of the intended CJK font. Fix: OMIT `name` and pass the
  // four attributes explicitly so eastAsia actually receives the CJK font.
  return new TextRun({
    text,
    font: { ascii: FONT_LATIN, hAnsi: FONT_LATIN, cs: FONT_LATIN, eastAsia: FONT_CJK },
    size: opts.size || SIZE_BODY,
    bold: opts.bold || false,
  });
}

// Mixed-run note paragraph that may contain Arabic terms. Pass the document
// `mode` so non-Arabic text uses the CJK font in ZH-AR notes (otherwise the
// note's Chinese characters fall back to a non-CJK font).
function noteParagraph(text, mode) {
  const arabicRe = /[\u0600-\u06FF\u0750-\u077F]+/g;
  const nonArabicFont = mode === "ZH-AR"
    ? { ascii: FONT_LATIN, hAnsi: FONT_LATIN, cs: FONT_LATIN, eastAsia: FONT_CJK }
    : FONT_LATIN;
  const runs = [];
  let last = 0;
  let m;
  while ((m = arabicRe.exec(text)) !== null) {
    if (m.index > last) {
      runs.push(new TextRun({ text: text.slice(last, m.index), font: nonArabicFont, size: SIZE_NOTE }));
    }
    runs.push(new TextRun({ text: m[0], font: FONT_ARABIC, size: SIZE_NOTE, rightToLeft: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) {
    runs.push(new TextRun({ text: text.slice(last), font: nonArabicFont, size: SIZE_NOTE }));
  }
  if (runs.length === 0) {
    runs.push(new TextRun({ text, font: nonArabicFont, size: SIZE_NOTE }));
  }
  return new Paragraph({ spacing: LINE_SPACING_1_2, children: runs });
}


function arabicRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT_ARABIC,
    size: opts.size || SIZE_BODY,
    bold: opts.bold || false,
    rightToLeft: true,
  });
}

// Left-column paragraph (English or Chinese, LTR)
function leftParagraph(text, mode) {
  return new Paragraph({
    spacing: LINE_SPACING_1_2,
    alignment: AlignmentType.LEFT,
    children: [mode === "ZH-AR" ? cjkRun(text) : latinRun(text)],
  });
}

// Right-column paragraph (Arabic, RTL)
function rightParagraph(text) {
  return new Paragraph({
    spacing: LINE_SPACING_1_2,
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    children: [arabicRun(text)],
  });
}

// (noteParagraph is defined above, after cjkRun, with mode-aware CJK font
//  handling. The earlier copy was removed to avoid a duplicate definition.)


// =============================================================================
// TABLE BUILDER (two columns: left text, right Arabic)
// =============================================================================

const CELL_BORDER = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
};

function headerCell(text, widthPct, isRtl) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" },
    borders: CELL_BORDER,
    children: [
      new Paragraph({
        spacing: LINE_SPACING_1_2,
        alignment: isRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
        bidirectional: isRtl,
        children: [
          new TextRun({
            text,
            font: isRtl ? FONT_ARABIC : FONT_LATIN,
            size: SIZE_BODY,
            bold: true,
            rightToLeft: isRtl,
          }),
        ],
      }),
    ],
  });
}

function bodyCell(paragraphs, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    borders: CELL_BORDER,
    children: paragraphs,
  });
}

function buildBilingualTable(doc) {
  const LEFT_WIDTH = 55;
  const RIGHT_WIDTH = 45;

  const headerLabels =
    doc.mode === "ZH-AR"
      ? { left: "中文", right: "العربية" }
      : { left: "English", right: "العربية" };

  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        headerCell(headerLabels.left, LEFT_WIDTH, false),
        headerCell(headerLabels.right, RIGHT_WIDTH, true),
      ],
    }),
    ...doc.rows.map(
      ([left, right]) =>
        new TableRow({
          children: [
            bodyCell([leftParagraph(left, doc.mode)], LEFT_WIDTH),
            bodyCell([rightParagraph(right)], RIGHT_WIDTH),
          ],
        })
    ),
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
  });
}

// =============================================================================
// DOCUMENT ASSEMBLY
// =============================================================================

function buildDocument(doc) {
  const header = new Header({
    children: [
      new Paragraph({
        spacing: LINE_SPACING_1_2,
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: doc.confidentiality,
            font: FONT_LATIN,
            size: 18,
            color: "808080",
          }),
        ],
      }),
    ],
  });

  const footer = new Footer({
    children: [
      new Paragraph({
        spacing: LINE_SPACING_1_2,
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Page ", font: FONT_LATIN, size: 18 }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT_LATIN, size: 18 }),
          new TextRun({ text: " of ", font: FONT_LATIN, size: 18 }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT_LATIN, size: 18 }),
        ],
      }),
    ],
  });

  const titleParagraph =
    doc.mode === "ZH-AR"
      ? new Paragraph({
          spacing: LINE_SPACING_1_2,
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_1,
          children: [cjkRun(doc.title, { size: SIZE_HEADING, bold: true })],
        })
      : new Paragraph({
          spacing: LINE_SPACING_1_2,
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_1,
          children: [latinRun(doc.title, { size: SIZE_HEADING, bold: true })],
        });

  const subtitleParagraph = new Paragraph({
    spacing: LINE_SPACING_1_2,
    alignment: AlignmentType.CENTER,
    children: [
      doc.mode === "ZH-AR"
        ? cjkRun(doc.subtitle, { size: SIZE_SUBTITLE, italics: true })
        : latinRun(doc.subtitle, { size: SIZE_SUBTITLE, italics: true }),
    ],
  });

  const spacer = new Paragraph({ spacing: LINE_SPACING_1_2, children: [new TextRun({ text: "" })] });

  const table = buildBilingualTable(doc);

  const noteHeader =
    doc.mode === "ZH-AR"
      ? new Paragraph({
          spacing: LINE_SPACING_1_2,
          heading: HeadingLevel.HEADING_2,
          children: [cjkRun("译注", { size: SIZE_BODY, bold: true })],
        })
      : new Paragraph({
          spacing: LINE_SPACING_1_2,
          heading: HeadingLevel.HEADING_2,
          children: [latinRun("Translator's Note", { size: SIZE_BODY, bold: true })],
        });

  const noteParagraphs = doc.translatorsNote.map((line) => noteParagraph(line, doc.mode));

  return new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: "portrait",
              // A4
              width: convertInchesToTwip(8.27),
              height: convertInchesToTwip(11.69),
            },
            margin: {
              top: convertInchesToTwip(1),
              bottom: convertInchesToTwip(1),
              left: convertInchesToTwip(1),
              right: convertInchesToTwip(1),
            },
          },
        },
        headers: { default: header },
        footers: { default: footer },
        children: [
          titleParagraph,
          subtitleParagraph,
          spacer,
          table,
          spacer,
          noteHeader,
          ...noteParagraphs,
        ],
      },
    ],
  });
}

// =============================================================================
// MAIN
// =============================================================================

async function main() {
  for (const docDef of DOCUMENTS) {
    const doc = buildDocument(docDef);
    const buffer = await Packer.toBuffer(doc);
    const outPath = path.join(process.cwd(), `${docDef.basename}.docx`);
    fs.writeFileSync(outPath, buffer);
    console.log(`Wrote ${outPath}`);
  }
}

main().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
