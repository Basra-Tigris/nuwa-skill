"""Prepare PDF and image sources for visual translation.

Usage:
  python scripts/prepare_visual_source.py source.pdf --out visual_source --dpi 220 --enhance
  python scripts/prepare_visual_source.py page-001.jpg page-002.png --out visual_source --enhance

The script writes normalized page images, manifest.json, and reading-notes.md.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".bmp"}


def import_pillow():
    try:
        from PIL import Image, ImageFilter, ImageOps

        return Image, ImageFilter, ImageOps
    except ImportError as exc:
        raise SystemExit(
            "Pillow is required for image normalization. Install with: "
            "python -m pip install pillow"
        ) from exc


def import_fitz():
    try:
        import fitz  # PyMuPDF

        return fitz
    except ImportError:
        return None


def is_pdf(path: Path) -> bool:
    return path.suffix.lower() == ".pdf"


def is_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def ensure_source(path: Path) -> None:
    if not path.exists():
        raise SystemExit(f"Source not found: {path}")
    if not path.is_file():
        raise SystemExit(f"Source is not a file: {path}")
    if not (is_pdf(path) or is_image(path)):
        supported = ", ".join(sorted([".pdf", *IMAGE_EXTENSIONS]))
        raise SystemExit(f"Unsupported source type: {path.suffix}. Supported: {supported}")


def normalize_image(src: Path, dest: Path, enhance: bool, quality: int) -> dict:
    Image, ImageFilter, ImageOps = import_pillow()
    with Image.open(src) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        if enhance:
            # Legal scans benefit from contrast and light sharpening. Keep color
            # when present so stamps, seals, and handwriting remain inspectable.
            if img.mode == "RGB":
                img = ImageOps.autocontrast(img)
            else:
                img = ImageOps.autocontrast(img.convert("L"))
            img = img.filter(ImageFilter.SHARPEN)
        if img.mode != "RGB":
            img = img.convert("RGB")
        dest.parent.mkdir(parents=True, exist_ok=True)
        img.save(dest, "JPEG", quality=quality, optimize=True)
        width, height = img.size
    return {"width": width, "height": height}


def render_pdf_with_pymupdf(src: Path, out_dir: Path, dpi: int) -> list[Path]:
    fitz = import_fitz()
    if fitz is None:
        return []
    rendered = []
    doc = fitz.open(src)
    try:
        matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            out_path = out_dir / f"render-{page_index + 1:04d}.png"
            pix.save(out_path)
            rendered.append(out_path)
    finally:
        doc.close()
    return rendered


def render_pdf_with_poppler(src: Path, out_dir: Path, dpi: int) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        return []
    prefix = out_dir / "render"
    cmd = [pdftoppm, "-png", "-r", str(dpi), str(src), str(prefix)]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise SystemExit(
            "pdftoppm failed:\n"
            + result.stdout.strip()
            + "\n"
            + result.stderr.strip()
        )
    def page_sort_key(path: Path) -> int:
        match = re.search(r"-(\d+)\.png$", path.name)
        return int(match.group(1)) if match else 0

    return sorted(out_dir.glob("render-*.png"), key=page_sort_key)


def render_pdf(src: Path, out_dir: Path, dpi: int) -> list[Path]:
    rendered = render_pdf_with_pymupdf(src, out_dir, dpi)
    if rendered:
        return rendered
    rendered = render_pdf_with_poppler(src, out_dir, dpi)
    if rendered:
        return rendered
    raise SystemExit(
        "Cannot render PDF: install PyMuPDF (python -m pip install pymupdf) "
        "or Poppler (pdftoppm)."
    )


def write_reading_notes(out_dir: Path, pages: list[dict]) -> None:
    lines = [
        "# Visual Reading Notes",
        "",
        "Use this file while reading the prepared pages. Record unclear text,",
        "numbers, stamps, signatures, and page defects before building the DOCX.",
        "",
    ]
    for page in pages:
        lines.extend(
            [
                f"## {page['page_id']}",
                "",
                f"- Source: {page['source']}",
                f"- Source page: {page.get('source_page') or 'image file'}",
                f"- Prepared image: {page['image']}",
                "- Source transcription:",
                "- English translation:",
                "- Chinese translation:",
                "- Uncertainties / defects:",
                "",
            ]
        )
    (out_dir / "reading-notes.md").write_text("\n".join(lines), encoding="utf-8")


def prepare_sources(args: argparse.Namespace) -> dict:
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    pages = []
    page_counter = 1
    source_paths = [Path(p).resolve() for p in args.sources]

    for source in source_paths:
        ensure_source(source)
        with tempfile.TemporaryDirectory(prefix="visual-source-") as tmp_name:
            tmp_dir = Path(tmp_name)
            if is_pdf(source):
                intermediate_pages = render_pdf(source, tmp_dir, args.dpi)
                source_type = "pdf"
            else:
                intermediate_pages = [source]
                source_type = "image"

            for source_page_number, intermediate in enumerate(intermediate_pages, start=1):
                page_id = f"page-{page_counter:03d}"
                dest = out_dir / f"{page_id}.jpg"
                dims = normalize_image(intermediate, dest, args.enhance, args.quality)
                pages.append(
                    {
                        "page_id": page_id,
                        "source": str(source),
                        "source_type": source_type,
                        "source_page": source_page_number if source_type == "pdf" else None,
                        "image": str(dest),
                        "width": dims["width"],
                        "height": dims["height"],
                    }
                )
                page_counter += 1

    manifest = {
        "created_by": "prepare_visual_source.py",
        "dpi": args.dpi,
        "enhance": args.enhance,
        "quality": args.quality,
        "page_count": len(pages),
        "pages": pages,
    }
    (out_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_reading_notes(out_dir, pages)
    return manifest


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Prepare PDFs and images as normalized page images for translation."
    )
    parser.add_argument("sources", nargs="+", help="PDF or image source files, in page order")
    parser.add_argument("--out", default="visual_source", help="Output directory")
    parser.add_argument("--dpi", type=int, default=220, help="PDF render DPI")
    parser.add_argument(
        "--enhance",
        action="store_true",
        help="Apply autocontrast and sharpening to prepared page images",
    )
    parser.add_argument("--quality", type=int, default=92, help="JPEG quality, 1-95")
    args = parser.parse_args(argv)
    if args.dpi < 72 or args.dpi > 600:
        raise SystemExit("--dpi must be between 72 and 600")
    if args.quality < 1 or args.quality > 95:
        raise SystemExit("--quality must be between 1 and 95")
    return args


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    manifest = prepare_sources(args)
    print(f"Wrote {manifest['page_count']} page image(s)")
    print(f"Manifest: {Path(args.out).resolve() / 'manifest.json'}")
    print(f"Reading notes: {Path(args.out).resolve() / 'reading-notes.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
