# PDF and Image Intake

Use this reference whenever the source is a PDF, scanned PDF, screenshot, page image, or photographed document. The goal is to turn unstable visual inputs into a stable page set before translation.

## Intake Rules

1. Preserve page order. For multiple images, use the order supplied by the user unless filenames clearly encode a safer page order.
2. Create one normalized page image per logical page. Do not merge pages unless the source itself is a two-page spread and the user wants it kept that way.
3. Keep a manifest that maps every output image back to its source file and page number.
4. Treat OCR as a draft aid. OCR text can seed transcription, but final translation must be checked visually against the page image.
5. Never infer missing text from context. If a stamp, signature, date, number, or clause is cropped, blurred, obscured, or too low-resolution to read, flag it in the translator's note.

## Standard Preparation

Run the helper from the skill directory or copy it to the working directory:

```powershell
& "<PYTHON_EXE>" scripts/prepare_visual_source.py "source.pdf" --out visual_source --dpi 220 --enhance
& "<PYTHON_EXE>" scripts/prepare_visual_source.py "page-001.jpg" "page-002.jpg" --out visual_source --enhance
```

Outputs:

- `visual_source/page-001.jpg`, `page-002.jpg`, etc. — normalized reading images
- `visual_source/manifest.json` — source-to-page mapping, dimensions, and processing flags
- `visual_source/reading-notes.md` — a page checklist for transcription, translation, and uncertainty tracking

For dense legal scans, rerun at `--dpi 300`. For photos with glare or colored stamps, try both with and without `--enhance`; enhancement can improve text but sometimes damages faint seals or handwriting.

## Reading Workflow

1. Open `manifest.json` and confirm page count and ordering.
2. Inspect each prepared page image. Rotate or recrop manually if the helper cannot make the page readable.
3. Build a working segmentation table:
   - page number
   - source segment or paragraph
   - English translation
   - Chinese translation
   - notes / uncertainty
4. Translate row by row into the bilingual DOCX template.
5. Carry page-level source defects into the translator's note.

For legal documents, give special attention to:

- headers, footers, file numbers, decision numbers, and article citations
- Arabic-Indic digits and mixed Latin identifiers
- stamps, signatures, license numbers, commercial registration numbers
- handwritten additions or marginal notes
- Hijri dates and any bracketed Gregorian conversions

## OCR Guidance

Use OCR only when it saves typing. Accept OCR output only after checking it against the image.

If Tesseract is available, Arabic OCR usually needs `ara`; mixed documents may need `ara+eng` or `ara+eng+chi_sim`. If the language data is not installed, skip OCR rather than installing large language packs during a translation task.

Common OCR failures:

- Arabic letters with dots confused or omitted
- Arabic-Indic digits changed to the wrong Western digit
- right-to-left line order reversed
- Latin identifiers like `M/75`, `DE-1447-00053`, or `GAC` split or reordered
- stamps and signatures hallucinated into body text

## Quality Bar

Before building the final Word files, verify that:

- every source page has a corresponding prepared page image
- page images are upright and readable at normal zoom
- no obvious page is missing, duplicated, or out of order
- every unreadable area is recorded
- every extracted/OCR number and identifier was visually checked

If the visual input is too poor for reliable translation, produce a partial working translation only for readable sections and say exactly what cannot be read.
