---
name: arabic-bilingual-translator
description: Translate Arabic, English, or Chinese source documents, PDFs, scanned PDFs, page images, screenshots, or photographed documents into bilingual Word files — one English-Arabic (EN left / AR right) and one Chinese-Arabic (ZH left / AR right) — preserving legal structure, reference identifiers, and party names, and surfacing rather than silently resolving ambiguities. This skill should be used when the user asks to translate an Arabic legal or general document, PDF, scan, image, or photo into English and/or Chinese (or vice versa: EN/ZH source into Arabic), produce a side-by-side bilingual Word file, render Arabic right-to-left correctly, read Arabic from PDFs/images, or handle Saudi/GCC legal terminology and Hijri dates. Triggers include phrases like "translate this Arabic PDF", "translate these images", "bilingual Arabic Word", "EN-AR translation", "ZH-AR translation", "Arabic legal translation", "沙特法律翻译", "阿拉伯语对照翻译", "阿语 PDF", "图片翻译".
---

# Arabic Bilingual Translator

## Overview

Turn an Arabic, English, or Chinese source document — including `.docx`, born-digital PDF, scanned PDF, screenshots, page images, or photographed documents — into two clean, professionally formatted Word files: one **English–Arabic** (EN left / AR right) and one **Chinese–Arabic** (ZH left / AR right), fit to circulate to counsel and stakeholders. The hard parts of this job are not the language — they are (1) reading scanned Arabic correctly, (2) preserving the structural and citational conventions a lawyer relies on, (3) rendering Arabic right-to-left in Word without breakage, and (4) surfacing rather than silently resolving the ambiguities that change legal meaning. This skill encodes the whole pipeline so each document does not have to be re-figured-out from scratch.

Two generation backends are bundled. Pick by environment:

- **docx-js** (`assets/build_bilingual_docx.js`) — preferred when Node.js is available; install `docx` locally into the managed workspace, never globally.
- **python-docx** (`assets/build_bilingual_py.py`) — fallback when only Python is available; install `python-docx` into the managed venv.

Run `python scripts/check_tools.py` once to see which backend, PDF/image intake tools, and visual-check tools are available on the current machine; it prints a ready-to-use command plan.

## Pipeline at a Glance

1. **Prepare and read the source correctly** — text extraction usually fails on scanned Arabic; normalize PDFs/images into page images and read visually.
2. **Detect source language and target pair** — AR source → emit EN-AR + ZH-AR; EN source → emit EN-AR (+ AR-ZH if asked); ZH source → emit ZH-AR (+ AR-ZH if asked). Default: always produce both EN-AR and ZH-AR when the source is Arabic; produce the matching pair when the source is EN or ZH unless the user asks for only one.
3. **Translate under the house rules** — preserve structure, references, and party names; never silently resolve ambiguity.
4. **Build the Word deliverables** — use the two-column table template; left column = EN or ZH translation, right column = Arabic source/translation; handle Arabic RTL.
5. **Validate and eyeball** — run the validator (docx-js path) or open-and-check (python-docx path), then render to image and inspect layout.
6. **Deliver with a translator's note** — list every judgment call and ambiguity for counsel to verify.

## Step 1 — Prepare and Read the Source Correctly

Arabic legal PDFs are very often scans or image-based exports. Naive extraction produces garbled, reordered, or empty output and silently corrupts numbers and names — unacceptable for a legal document. Do not trust `pdftotext`, OCR, or model-read text on an Arabic PDF/image without visual checking.

**Accepted source inputs:**

- `.docx` or pasted text — extract text directly, then spot-check structure and references if an original rendering is available.
- Born-digital `.pdf` — try text extraction, but validate against rendered pages because bidirectional Arabic often reorders numbers and punctuation.
- Scanned `.pdf` — render pages to images first; read visually or use OCR only as a draft aid.
- Image files (`.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.webp`, `.bmp`) — normalize orientation and quality first; treat every image as one page unless the user says otherwise.
- Multiple page images — preserve the filename/order supplied by the user and flag missing/duplicate pages before translating.

For PDF or image sources, read `references/pdf-image-intake.md`, then run the intake helper to create a stable page set and manifest:

```powershell
& "<PYTHON_EXE>" scripts/prepare_visual_source.py "source.pdf" --out visual_source --dpi 220 --enhance
& "<PYTHON_EXE>" scripts/prepare_visual_source.py "page-001.jpg" "page-002.jpg" --out visual_source --enhance
```

The helper writes normalized page images plus `manifest.json` and `reading-notes.md`. Use the manifest to keep page order, source page numbers, and any unreadable areas traceable through the final translator's note.

```bash
# First, see what the file actually is and whether text extraction is usable
pdfinfo source.pdf
pdftotext -layout source.pdf - | head -40    # if garbled / empty / scrambled, DO NOT use it
```

When extraction is unreliable (the common case), rasterize the pages and read them visually — this is the reliable path for Arabic. The helper above uses PyMuPDF when available and falls back to Poppler; direct Poppler usage is:

```bash
pdftoppm -jpeg -r 200 source.pdf page        # 200 dpi is the sweet spot for Arabic legibility
ls page*.jpg
```

Then view each `page-N.jpg` or prepared page image and read the Arabic directly from the image. Resolution matters: below ~150 dpi Arabic diacritics and digits blur; 200-220 dpi is reliable, go to 300 for dense small print.

**Windows fallback when Poppler is missing:** use Python with `pdf2image` + `Pillow`, or `pymupdf` (fitz), installed into the managed venv:

```powershell
& "<PYTHON_EXE>" -m pip install pdf2image pillow pymupdf
```

(`pdf2image` still needs Poppler installed on Windows; `pymupdf` does not and is the recommended no-Poppler path.)

For `.docx` or born-digital Arabic PDFs where extraction is clean, use the text directly — but still spot-check digits, reference numbers, and Hijri dates against the rendered page, because bidirectional text frequently reorders numbers during extraction. For image inputs, do not paraphrase from visual impression alone: segment the source into rows/paragraphs, transcribe or summarize each source segment, then translate segment-by-segment so omissions are visible.

If the user pasted Arabic text directly into chat, work from that, but flag that the translation is of the text as given and cannot be verified against an original. If the user supplies only low-resolution, cropped, rotated, blurred, or glare-obscured images, translate readable text and list the unreadable regions in the translator's note instead of guessing.

## Step 2 — Detect Source Language and Target Pair

Determine the source language automatically from script and content:

- Contains Arabic script → source is **AR**. Default output: both **EN-AR** and **ZH-AR** bilingual files (Arabic on the right in each).
- Contains Chinese characters (and no Arabic) → source is **ZH**. Default output: **ZH-AR** bilingual (ZH left / AR right). Add **AR-ZH** with Arabic on the left only if the user asks.
- Otherwise (Latin script) → source is **EN**. Default output: **EN-AR** bilingual (EN left / AR right). Add **AR-EN** with Arabic on the left only if the user asks.

When the user explicitly asks for only one of the two default outputs, honor that.

**Mode detection (legal vs. general):** scan the source for legal-domain markers — Saudi/GCC authority names (GAC, SDAIA, BOG), statute citations (M/75, Royal Decree), case/decision numbers (Decision No. DE-…), Hijri dates, formal salutations to ministers or courts, or competition-law terminology. If two or more markers appear, run in **legal mode** (load `references/glossary-legal.md`, apply Hijri-date rules, preserve argument hierarchy). Otherwise run in **general mode** (load `references/glossary-general.md`, lighter structure preservation). When unsure, default to legal mode for safety — it is stricter and a superset of general rules.

## Step 3 — Translate Under the House Rules

These rules exist because in a legal document a small infidelity (a flipped number, a quietly "fixed" ambiguity, an inconsistent party name) is a substantive error, not a stylistic one.

**Preserve structure exactly.** Keep the document's numbered/lettered argument hierarchy, headings, bold emphasis, salutation, and signature block in the same order and nesting as the original. A Saudi reply memorandum's numbered sections and a decision's operative holdings carry legal weight by their structure.

**Preserve all reference identifiers verbatim.** Case numbers, decision numbers (e.g. Decision No. DE-1447-00053), commercial registration numbers, license numbers, statute and article citations, and Hijri dates are copied exactly, never normalized or "corrected." Keep Hijri dates as written; add the Gregorian equivalent in square brackets, e.g. `7/8/1447 AH [≈ 26 January 2026]` — and mark any conversion as approximate in the translator's note.

**Translate party and authority names consistently.** Pick one rendering of each entity, person, court, and authority and use it throughout. See `references/glossary-legal.md` for standard renderings of GAC, SDAIA, BOG, M/75, etc. Transliterate personal names consistently; on first mention give the Arabic in parentheses.

**Retain key legal terms bilingually when the target is Chinese.** For AR→ZH output, write the Chinese narrative but keep the controlling legal term in its Arabic and/or English form in parentheses on first use — e.g. `滥用市场支配地位（إساءة استخدام الوضع المهيمن / abuse of dominant position）`. This matches the established house style and lets a reviewer cross-check against the source and against English filings.

**Match register and formality.** Saudi administrative and court Arabic uses formal salutations, honorifics, and polite-imperative constructions. Render these into natural, formal legal English/Chinese rather than word-for-word calques, but do not flatten a binding directive into a mere request.

**Handle corporate gender.** Arabic assigns grammatical gender (companies are often feminine); render as gender-neutral in English. Note this once if it could affect reading.

**NEVER silently resolve an ambiguity.** This is the cardinal rule. If a word is genuinely ambiguous, if the scan is unclear, if a number looks internally inconsistent, or if a term could be read two ways with different legal consequences (the classic example: `أعلى` "higher" vs. a context implying "lower" in a below-cost-pricing argument), translate the most defensible reading and flag it in the translator's note (Step 6). Do not pick silently and do not omit. A flagged uncertainty is useful to counsel; a hidden one is a liability.

**Mark inline only when unavoidable.** Prefer end-of-document notes. If an inline marker is truly needed, use `[Translator's note: …]` so it can never be mistaken for source text.

## Step 4 — Build the Word Deliverables

Two output files are produced by default (one when the source is EN or ZH and the user did not ask for both directions):

- `<basename>_EN-AR.docx` — left column English, right column Arabic
- `<basename>_ZH-AR.docx` — left column Chinese, right column Arabic

Use the matching template as the starting point:

- **docx-js path:** copy `assets/build_bilingual_docx.js` to the working directory and edit the content arrays. It is preconfigured for legal output: A4, formal serif body (Times New Roman for Latin, Amiri for any retained Arabic, SimSun/宋体 for Chinese), justified text, numbered-section support, two-column table support, a header/footer with page numbers and a draft/confidentiality line, and **1.2× line spacing on every paragraph** (see Line Spacing below).
- **python-docx path:** copy `assets/build_bilingual_py.py` to the working directory and edit the content arrays. Same formatting target as above (including 1.2× line spacing); RTL is set via the `w:bidi`/`w:rtl` XML properties because python-docx has no high-level RTL API.

### Layout: Two-Column Side-by-Side Table

Both EN-AR and ZH-AR files use a **two-column table** with the source/translation pair aligned row by row:

- **Left column** (≈55% width): English (in EN-AR file) or Chinese (in ZH-AR file). LTR paragraph, left-aligned, justified.
- **Right column** (≈45% width): Arabic. RTL paragraph, right-aligned, `bidirectional: true` on the paragraph and `rightToLeft: true` + Arabic font on every Arabic run.

This matches the user's standing requirement: **English and Chinese on the left, Arabic on the right.**

### Line Spacing (mandatory)

Every paragraph in the output — title, subtitle, table header cells, table body cells (both LTR and RTL), translator's note heading, and every note line — uses **1.2× line spacing**. This is a hard formatting rule for this skill; do not ship a file with default single spacing.

**docx-js implementation:** a single constant is applied to every `Paragraph` via the `spacing` option:

```javascript
const LINE_SPACING_1_2 = { line: 288, lineRule: LineRuleType.AUTO };
// 288 = 240 * 1.2; when lineRule = AUTO, `line` is in 240ths of a line.

new Paragraph({ spacing: LINE_SPACING_1_2, /* ...rest */ });
```

**python-docx implementation:** a helper sets the multiple-spacing rule on the paragraph format:

```python
from docx.enum.text import WD_LINE_SPACING

def set_1_2_spacing(paragraph):
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.2

# Call on every paragraph after creation.
```

If you extend the templates with new paragraph types (e.g. a footnote section, an annex), apply the same spacing — do not let any paragraph fall back to single spacing.

### Arabic Right-to-Left Handling

Whenever Arabic text is retained in the output (always, in this skill), it must render RTL or it will display backwards/broken.

**docx-js:**

```javascript
// Arabic paragraph in the right column
new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  children: [ new TextRun({ text: "النص العربي", rightToLeft: true, font: "Amiri", size: 24 }) ]
})

// Mixed run inside otherwise-LTR text (e.g. an Arabic legal term inside English)
new Paragraph({ children: [
  new TextRun({ text: "the term ", font: "Times New Roman" }),
  new TextRun({ text: "الوضع المهيمن", rightToLeft: true, font: "Amiri" }),
  new TextRun({ text: " (dominant position)", font: "Times New Roman" })
]})
```

For side-by-side tables, set the Arabic cell's paragraph `alignment: AlignmentType.RIGHT` and `bidirectional: true`. Follow the table rules: dual column widths, `ShadingType.CLEAR` for header row, explicit cell margins, no row spanning across pages when avoidable.

**python-docx (RTL via raw XML, since the library exposes no high-level API):**

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_rtl_paragraph(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def set_rtl_run(run):
    rPr = run._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
```

### Other Formatting Carried From the Source

Reproduce comparative data tables (e.g. an authority's figures vs. a transport-authority order), bolded emphases, and the signature/attestation block (attorney name, license number, capacity). Keep statute and decision citations in a consistent typographic treatment so they are easy to scan.

## Step 5 — Validate and Eyeball

Never hand over an unvalidated or unseen file.

**docx-js path:**

```bash
node build_bilingual_docx.js
# render to image and actually look at the layout — RTL bugs and table breakage are visual
# Windows: use LibreOffice if installed
soffice --headless --convert-to pdf output.docx
# or via Python (pymupdf) to render PDF pages to JPG for inspection
```

**python-docx path:**

```powershell
& "<PYTHON_EXE>" build_bilingual_py.py
# open the .docx in Word/LibreOffice and visually check, or convert to PDF then to images
```

Then view the rendered pages. Confirm:

- Arabic reads right-to-left and is **not reversed** (the most common bug)
- Numbers and reference IDs match the source
- Tables are intact; column widths look right
- The signature block is present
- Nothing overflows the margins
- The EN-AR file has English on the left, Arabic on the right
- The ZH-AR file has Chinese on the left, Arabic on the right

Fix and re-render before delivering. Copy the final files to the project's working directory (or `outputs/` if the user asks) and present them.

## Step 6 — Deliver With a Translator's Note

Every delivery ends with a short, explicit note to counsel. This is what makes the output professionally usable rather than just a draft. Keep it tight and concrete.

**Always include, after presenting the file:**

> **Translator's note** (for counsel verification — this is a working translation, not a certified/sworn translation):
>
> 1. **Judgment calls** — terms rendered one way among defensible options, with the alternative:
>    - "`<term>`" rendered as "`<chosen>`"; could also read "`<alternative>`", which would affect `<consequence>`. **[Risk: High/Medium/Low]**
> 2. **Ambiguities / source defects** — anything unclear in the scan, internally inconsistent, or potentially mis-stated in the source. **[Risk: …]**
> 3. **Date conversions** — any Hijri→Gregorian conversions, flagged as approximate.
> 4. **Terminology choices** the firm may wish to standardize (entity names, court names, statute titles).
> 5. **Backend used** — docx-js or python-docx; note any tooling that was unavailable and how it was worked around.

If there were no judgment calls or ambiguities, say so explicitly rather than omitting the note — counsel needs to know the document was clean.

Do not overstate certainty. This is a working translation to support a lawyer's review; it is not a certified or sworn translation, and the note should say so.

## Environment and Dependencies

This skill is **self-contained** — it does not depend on any external "docx skill." All generation logic lives in `assets/`.

**Required (always):**
- Node.js 22+ **or** Python 3.13+ (use the bundled runtime paths available in the current Codex session)
- `docx` npm package (docx-js path) — prefer the bundled Node package directory when it already contains `docx`; otherwise install locally into the working directory, never globally:
  ```powershell
  & "<PNPM_EXE>" add docx
  # then run with NODE_PATH set:
  $env:NODE_PATH = "<NODE_MODULES>"
  & "<NODE_EXE>" build_bilingual_docx.js
  ```
- `python-docx` (python-docx path) — install into the active managed Python if missing:
  ```powershell
  & "<PYTHON_EXE>" -m pip install python-docx
  ```

**Optional (PDF source reading & visual check):**
- Pillow (`PIL`) — for normalizing page images, screenshots, and photographed documents
- `pymupdf` (Python) — Poppler-free PDF text extraction and page rendering; preferred for `scripts/prepare_visual_source.py`
- Poppler (`pdftoppm`, `pdftotext`, `pdfinfo`) — fallback for rasterizing and extracting Arabic PDFs
- Tesseract OCR with Arabic/English/Chinese language data — optional draft text aid only; always visually verify
- LibreOffice (`soffice`) — for converting the output `.docx` to PDF for the visual check

Run `python scripts/check_tools.py` to see which are available and get a tailored command plan.

## Reference Files

- `references/glossary-legal.md` — standard English/Chinese renderings of recurring Saudi/GCC legal entities, courts, statutes, and competition-law terms, plus Hijri-date guidance. Read it before translating in legal mode so entity and term renderings are consistent across documents.
- `references/glossary-general.md` — high-frequency general-domain Arabic terms with EN/ZH renderings. Read it before translating in general mode.
- `references/pdf-image-intake.md` — PDF/image intake workflow, OCR rules, page-order checks, and visual-reading quality controls. Read it whenever the source is a PDF, scanned PDF, screenshot, page image, or photo.
- `references/rtl-and-dates.md` — detailed RTL rendering rules for both docx-js and python-docx, common RTL bugs and how to spot them, and the Hijri→Gregorian conversion recipe with worked examples.

## Asset Files

- `assets/build_bilingual_docx.js` — docx-js build template (A4, RTL-aware, legal formatting, two-column side-by-side). Copy and edit; do not edit in place.
- `assets/build_bilingual_py.py` — python-docx build template (same formatting target, RTL via raw XML). Copy and edit; do not edit in place.
- `scripts/check_tools.py` — detect available tools and print a command plan.
- `scripts/prepare_visual_source.py` — convert PDF pages and image files into normalized per-page images plus a manifest for visual translation.
