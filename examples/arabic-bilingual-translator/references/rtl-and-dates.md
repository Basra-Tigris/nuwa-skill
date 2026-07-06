# RTL Rendering and Hijri Dates

Detailed rules for rendering Arabic right-to-left in Word output (both docx-js and python-docx), common RTL bugs and how to spot them, and the Hijri→Gregorian conversion recipe with worked examples.

## Why RTL Handling Matters

Arabic is a right-to-left script. If you put Arabic text into a Word paragraph without setting RTL properties, the text will display in one of these broken ways:

1. **Reversed character order** — the whole string appears backwards (left-to-right).
2. **Reversed word order** — characters within each word are correct, but words flow left-to-right.
3. **Scrambled numbers and Latin terms** — Arabic-Indic digits, Western digits, and embedded Latin terms (e.g. `GAC`, `M/75`) reorder incorrectly because the bidirectional algorithm mis-classifies them.
4. **Punctuation drift** — periods, commas, parentheses, and brackets end up on the wrong side of the word they belong to.

The fix is to set **both** the paragraph-level `bidirectional` (bidi) flag and the run-level `rightToLeft` (rtl) flag. Setting only one of them leaves the text half-broken.

## docx-js (preferred when Node.js available)

### Paragraph-level RTL

```javascript
const { Paragraph, AlignmentType, TextRun } = require("docx");

new Paragraph({
  bidirectional: true,            // paragraph flows right-to-left
  alignment: AlignmentType.RIGHT, // text right-aligned
  children: [
    new TextRun({
      text: "النص العربي",
      rightToLeft: true,          // run is RTL
      font: "Amiri",              // Arabic-capable font
      size: 24,                   // half-points; 24 = 12pt
    }),
  ],
})
```

### Mixed Run (Arabic term inside LTR paragraph)

```javascript
new Paragraph({ children: [
  new TextRun({ text: "the term ", font: "Times New Roman" }),
  new TextRun({ text: "الوضع المهيمن", rightToLeft: true, font: "Amiri" }),
  new TextRun({ text: " (dominant position)", font: "Times New Roman" }),
]})
```

The mixed-run case is what makes bilingual Chinese output work: the paragraph is LTR Chinese, but each Arabic legal term in parentheses is its own RTL run with `rightToLeft: true` and `font: "Amiri"`.

### Arabic Cell in a Two-Column Table

```javascript
const { TableCell, WidthType, ShadingType } = require("docx");

new TableCell({
  width: { size: 45, type: WidthType.PERCENTAGE },
  children: [
    new Paragraph({
      bidirectional: true,
      alignment: AlignmentType.RIGHT,
      children: [ new TextRun({ text: "النص العربي", rightToLeft: true, font: "Amiri", size: 24 }) ],
    }),
  ],
})
```

## python-docx (fallback)

python-docx has **no high-level API** for RTL. Set the underlying OOXML elements directly:

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_rtl_paragraph(paragraph):
    """Mark a paragraph as right-to-left (bidi)."""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def set_rtl_run(run):
    """Mark a run as right-to-left (rtl)."""
    rPr = run._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)

def set_arabic_font(run, font_name="Amiri"):
    """Set both ascii and complex-script font for Arabic runs."""
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:cs'), font_name)   # complex script font
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
```

Usage:

```python
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

doc = Document()
p = doc.add_paragraph()
p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
set_rtl_paragraph(p)
run = p.add_run("النص العربي")
set_rtl_run(run)
set_arabic_font(run, "Amiri")
```

For table cells, do the same on each paragraph inside the cell.

## Common RTL Bugs and How to Spot Them

After rendering the output to PDF and then to images (see Step 5 of SKILL.md), check each Arabic page for these failure modes:

| Bug | What it looks like | Cause | Fix |
|---|---|---|---|
| Whole string reversed | Arabic reads LTR, characters backwards | `bidirectional` not set on paragraph | Add `bidirectional: true` (docx-js) or `w:bidi` (python-docx) |
| Words correct, order wrong | Each word reads correctly but flows LTR | `rightToLeft` not set on run | Add `rightToLeft: true` (docx-js) or `w:rtl` (python-docx) |
| Numbers wrong | `DE-1447-00053` shows as `3500-7441-ED` or similar | Number reordering under bidi algorithm | Wrap the number in its own LTR run with explicit font |
| Punctuation drift | `النص العربي.` shows period to the left of the word | Punctuation inherits paragraph direction | Put the period inside the same RTL run as the word |
| Latin term inside Arabic misplaced | `GAC` appears at the wrong position in the Arabic sentence | The Latin term is its own LTR run inside an RTL paragraph | This is actually correct behavior — verify it reads in logical order |
| Parentheses flipped | `(dominant position)` shows as `)dominant position(` | Parentheses direction follows paragraph bidi | Wrap the parenthesized phrase as a single LTR run when it is English text |

## Hijri → Gregorian Conversion

The Hijri (Islamic) calendar is lunar, ~354 or 355 days per year, so it drifts ~11 days per solar year against the Gregorian calendar. Saudi official documents use the **Umm al-Qura** calendar by convention.

### Conversion recipe

1. Identify the Hijri date components: day, month, year AH.
2. Convert using a reliable library — do **not** hand-compute, off-by-one errors are common.
   - Python: `pip install hijri-converter` (Umm al-Qura), or `hijridate`
   - Online: use a trusted converter such as the Umm al-Qura converter
3. Format the Gregorian equivalent as `DD Month YYYY`.
4. Insert into the document as `<Hijri as written> AH [≈ <Gregorian>]`.
5. Record every conversion in the translator's note as approximate.

### Python example

```python
# pip install hijri-converter
from hijri_converter import convert

g = convert.Hijri(1447, 8, 7).to_gregorian()
print(f"{g.year}-{g.month:02d}-{g.day:02d}")  # 2026-01-26 (approximate)
```

### Worked example

Source: `7/8/1447 AH`

- Hijri: 7 Sha'ban 1447 AH
- Gregorian: ≈ 26 January 2026
- Output in document: `7/8/1447 AH [≈ 26 January 2026]`
- Translator's note entry: "Hijri date `7/8/1447 AH` converted to Gregorian as ≈ 26 January 2026 (Umm al-Qura, approximate)."

### Hijri month names (for reference)

| # | Arabic | Transliteration |
|---|---|---|
| 1 | محرّم | Muharram |
| 2 | صفر | Safar |
| 3 | ربيع الأول | Rabi' al-Awwal |
| 4 | ربيع الثاني | Rabi' al-Thani |
| 5 | جمادى الأولى | Jumada al-Awwal |
| 6 | جمادى الثانية | Jumada al-Thani |
| 7 | رجب | Rajab |
| 8 | شعبان | Sha'ban |
| 9 | رمضان | Ramadan |
| 10 | شوّال | Shawwal |
| 11 | ذو القعدة | Dhu al-Qi'dah |
| 12 | ذو الحجّة | Dhu al-Hijjah |

### Important caveats

- **Always mark the conversion as approximate** with `≈` or the word "approximately". Hijri days begin at sunset, and the printed date can shift by ±1 day relative to the Gregorian calendar depending on moon sighting.
- **Never "correct" the Hijri date in the source.** If the source says `7/8/1447 AH`, write `7/8/1447 AH` verbatim in the Arabic column and in the document body. The Gregorian bracket is a courtesy annotation, not a correction.
- **If the Hijri date is internally inconsistent** (e.g. the day number is out of range, or two different Hijri dates appear for the same event), translate as written and flag in the translator's note under "Ambiguities / source defects".
