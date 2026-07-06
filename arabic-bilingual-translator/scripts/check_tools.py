"""check_tools.py
Detect which generation backends and PDF/image tools are available on this machine,
and print a tailored command plan for translating an Arabic document.

Usage:
  python check_tools.py
"""

import os
import shutil
import subprocess
import sys

# Prefer the current interpreter and the Codex Desktop bundled runtime.
# Override with PYTHON_EXE, NODE_EXE, NODE_MODULES, PNPM_EXE if needed.
HOME = os.path.expanduser("~")
CODEX_DEPS = os.path.join(
    HOME,
    ".cache",
    "codex-runtimes",
    "codex-primary-runtime",
    "dependencies",
)
BUNDLED_PYTHON = os.path.join(CODEX_DEPS, "python", "python.exe")
BUNDLED_NODE = os.path.join(CODEX_DEPS, "node", "bin", "node.exe")
BUNDLED_NODE_MODULES = os.path.join(CODEX_DEPS, "node", "node_modules")
BUNDLED_PNPM = os.path.join(CODEX_DEPS, "bin", "pnpm.cmd")

PYTHON_EXE = os.environ.get("PYTHON_EXE") or (
    BUNDLED_PYTHON if os.path.exists(BUNDLED_PYTHON) else sys.executable
)
NODE_EXE = os.environ.get("NODE_EXE") or (
    BUNDLED_NODE if os.path.exists(BUNDLED_NODE) else (shutil.which("node") or "node")
)
NODE_MODULES = os.environ.get("NODE_MODULES") or os.environ.get("NODE_PATH") or BUNDLED_NODE_MODULES
PNPM_EXE = os.environ.get("PNPM_EXE") or (
    BUNDLED_PNPM if os.path.exists(BUNDLED_PNPM) else (shutil.which("pnpm") or "pnpm")
)


def have(cmd):
    """Return True if cmd is on PATH."""
    return shutil.which(cmd) is not None


def have_module(python_exe, module):
    """Return True if a Python module is importable in the given interpreter."""
    try:
        r = subprocess.run(
            [python_exe, "-c", f"import {module}; print('{module} OK')"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return r.returncode == 0 and "OK" in r.stdout
    except Exception:
        return False


def have_npm_package(package, module_paths):
    """Return True if a node package is installed in one of the given module paths."""
    for module_path in module_paths.split(os.pathsep):
        if module_path and os.path.isdir(os.path.join(module_path, package)):
            return True
    return False


def main():
    print("=" * 70)
    print("Arabic Bilingual Translator - Environment Check")
    print("=" * 70)
    print()

    # --- Runtimes ---
    print("[Runtimes]")
    py_ok = os.path.exists(PYTHON_EXE)
    node_ok = os.path.exists(NODE_EXE)
    print(f"  Python (managed)      {PYTHON_EXE}  -> {'OK' if py_ok else 'MISSING'}")
    print(f"  Node.js (managed)     {NODE_EXE}  -> {'OK' if node_ok else 'MISSING'}")
    print(f"  Node modules          {NODE_MODULES}  -> {'OK' if os.path.exists(NODE_MODULES) else 'MISSING'}")
    print()

    # --- Generation backends ---
    print("[Generation backends]")
    pydocx_ok = have_module(PYTHON_EXE, "docx") if py_ok else False
    docx_js_ok = have_npm_package("docx", NODE_MODULES) if node_ok else False
    print(f"  python-docx           -> {'OK' if pydocx_ok else 'NOT INSTALLED'}")
    print(f"  docx-js (npm)         -> {'OK' if docx_js_ok else 'NOT INSTALLED'}")
    print()

    # --- PDF/image tools ---
    print("[PDF/image source reading & visual check]")
    poppler_ok = have("pdftoppm") and have("pdftotext") and have("pdfinfo")
    soffice_ok = have("soffice") or have("libreoffice")
    pymupdf_ok = have_module(PYTHON_EXE, "fitz") if py_ok else False
    pillow_ok = have_module(PYTHON_EXE, "PIL") if py_ok else False
    tesseract_ok = have("tesseract")
    print(f"  Pillow (PIL)          -> {'OK' if pillow_ok else 'NOT INSTALLED'}")
    print(f"  Poppler (pdftoppm)    -> {'OK' if poppler_ok else 'NOT INSTALLED'}")
    print(f"  pymupdf (fitz)        -> {'OK' if pymupdf_ok else 'NOT INSTALLED'}")
    print(f"  Tesseract OCR         -> {'OK' if tesseract_ok else 'NOT INSTALLED'}")
    print(f"  LibreOffice (soffice) -> {'OK' if soffice_ok else 'NOT INSTALLED'}")
    print()

    # --- Recommended backend ---
    print("[Recommended backend]")
    if docx_js_ok:
        backend = "docx-js"
        print(f"  -> {backend} (preferred; full RTL support via docx-js API)")
    elif pydocx_ok:
        backend = "python-docx"
        print(f"  -> {backend} (Node/docx-js not available; RTL via raw XML)")
    elif node_ok:
        backend = "docx-js (after install)"
        print(f"  -> {backend} (Node available, install docx package)")
    elif py_ok:
        backend = "python-docx (after install)"
        print(f"  -> {backend} (Python available, install python-docx)")
    else:
        backend = "NONE"
        print("  -> NO BACKEND AVAILABLE (install Python or Node)")
    print()

    # --- Command plan ---
    print("[Command plan]")
    if backend.startswith("docx-js"):
        if not docx_js_ok:
            print("  # 1. Install docx package locally (NEVER globally):")
            print('  pnpm add docx')
            print()
        print("  # 2. Copy the build template and edit content:")
        print('  cp "<skill-dir>/assets/build_bilingual_docx.js" .')
        print("  # edit DOCUMENTS array in build_bilingual_docx.js")
        print()
        print("  # 3. Build the .docx files:")
        print(f'  $env:NODE_PATH = "{NODE_MODULES}"')
        print(f'  & "{NODE_EXE}" build_bilingual_docx.js')
        print()
    elif backend.startswith("python-docx"):
        if not pydocx_ok:
            print("  # 1. Install python-docx into the active managed Python:")
            print(f'  & "{PYTHON_EXE}" -m pip install python-docx')
            print()
        print("  # 2. Copy the build template and edit content:")
        print('  cp "<skill-dir>/assets/build_bilingual_py.py" .')
        print("  # edit DOCUMENTS list in build_bilingual_py.py")
        print()
        print("  # 3. Build the .docx files:")
        print(f'  & "{PYTHON_EXE}" build_bilingual_py.py')
        print()

    print("  # Prepare PDF/image sources for visual translation:")
    print(
        f'  & "{PYTHON_EXE}" scripts/prepare_visual_source.py "source.pdf" '
        "--out visual_source --dpi 220 --enhance"
    )
    print(
        f'  & "{PYTHON_EXE}" scripts/prepare_visual_source.py "page-001.jpg" '
        '"page-002.jpg" --out visual_source --enhance'
    )
    print()

    if not pillow_ok:
        print("  # Image preparation requires Pillow:")
        print(f'  & "{PYTHON_EXE}" -m pip install pillow')
        print()

    if not poppler_ok and not pymupdf_ok:
        print("  # PDF source reading: neither Poppler nor pymupdf available.")
        print("  #   Option A (preferred): install pymupdf (no system deps):")
        print(f'  & "{PYTHON_EXE}" -m pip install pymupdf')
        print("  #   Option B: install Poppler for Windows from")
        print("  #     https://github.com/oschwartz10612/poppler-windows/releases")
        print("  #     and add its bin/ to PATH.")
        print()

    if not tesseract_ok:
        print("  # OCR: Tesseract is optional. Use it only as a draft aid;")
        print("  #      always visually verify OCR text against the page image.")
        print()

    if not soffice_ok:
        print("  # Visual check: LibreOffice not on PATH.")
        print("  #   Install from https://www.libreoffice.org/download/ for PDF conversion.")
        print("  #   Or open the .docx in Word and visually verify RTL rendering.")
        print()

    print("=" * 70)
    print("Done. Use the plan above to set up the environment.")
    print("=" * 70)


if __name__ == "__main__":
    main()
