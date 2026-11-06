"""
pdf_to_images.py
Convert PDF documents into high-resolution images (PNG/JPG) for multimodal AI inspection.
Supports page selection, custom DPI, and uses PyMuPDF (fast) with ImageMagick fallback.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Set


def parse_page_selection(pages_str: Optional[str], total_pages: int) -> Set[int]:
    """
    Parse a page range string like '1,3,5-8' into a set of 1-based page numbers.
    If pages_str is None or empty, returns all pages 1..total_pages.
    """
    if not pages_str or pages_str.strip().lower() == "all":
        return set(range(1, total_pages + 1))

    selected: Set[int] = set()
    parts = pages_str.split(",")
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            match = re.match(r"^(\d+)-(\d+)$", part)
            if match:
                start, end = int(match.group(1)), int(match.group(2))
                for p in range(max(1, start), min(total_pages, end) + 1):
                    selected.add(p)
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    selected.add(p)
            except ValueError:
                pass
    return selected


def render_pdf_with_pymupdf(
    pdf_path: Path,
    out_dir: Path,
    pages_to_render: Set[int],
    dpi: int = 150,
    fmt: str = "png",
    prefix: str = "page",
) -> List[Path]:
    """Render PDF pages to images using PyMuPDF (fitz)."""
    try:
        import pymupdf  # PyMuPDF
    except ImportError:
        import fitz as pymupdf

    doc = pymupdf.open(str(pdf_path.resolve()))
    out_dir.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    rendered_files: List[Path] = []

    for page_num in sorted(pages_to_render):
        if page_num < 1 or page_num > len(doc):
            continue
        page = doc[page_num - 1]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_file = out_dir / f"{prefix}_{page_num:03d}.{fmt}"
        pix.save(str(out_file))
        rendered_files.append(out_file)

    doc.close()
    return rendered_files


def render_pdf_with_magick(
    pdf_path: Path,
    out_dir: Path,
    pages_to_render: Set[int],
    dpi: int = 150,
    fmt: str = "png",
    prefix: str = "page",
) -> List[Path]:
    """Fallback to ImageMagick if PyMuPDF is not installed."""
    magick = shutil.which("magick")
    if not magick:
        return []

    out_dir.mkdir(parents=True, exist_ok=True)
    rendered_files: List[Path] = []

    for page_num in sorted(pages_to_render):
        out_file = out_dir / f"{prefix}_{page_num:03d}.{fmt}"
        # ImageMagick uses 0-based page index in brackets: input.pdf[0]
        cmd = [
            magick,
            "-density",
            str(dpi),
            f"{pdf_path.resolve()}[{page_num - 1}]",
            "-background",
            "white",
            "-alpha",
            "remove",
            "-alpha",
            "off",
            str(out_file.resolve()),
        ]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if out_file.exists() and out_file.stat().st_size > 0:
                rendered_files.append(out_file)
            else:
                # If page rendering failed (e.g. page out of range), do not loop indefinitely
                if len(rendered_files) > 0:
                    break
        except Exception:
            break

    return rendered_files


def pdf_to_images(
    pdf_path: str | Path,
    output_dir: Optional[str | Path] = None,
    pages: Optional[str] = None,
    dpi: int = 150,
    fmt: str = "png",
    prefix: Optional[str] = None,
) -> List[Path]:
    """
    Convert a PDF file to a list of image paths.
    """
    pdf = Path(pdf_path).resolve()
    if not pdf.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf}")

    out_d = Path(output_dir).resolve() if output_dir else pdf.parent / f"{pdf.stem}_images"
    pref = prefix or pdf.stem

    # Determine total pages
    total_pages = 0
    try:
        import pymupdf
        with pymupdf.open(str(pdf)) as doc:
            total_pages = len(doc)
    except Exception:
        # Fallback estimation if PyMuPDF unavailable
        total_pages = 100

    selected_pages = parse_page_selection(pages, total_pages)

    # Try PyMuPDF first
    try:
        results = render_pdf_with_pymupdf(pdf, out_d, selected_pages, dpi=dpi, fmt=fmt, prefix=pref)
        if results:
            return results
    except Exception as e:
        print(f"[PyMuPDF] Failed: {e}, attempting ImageMagick fallback...", file=sys.stderr)

    # Try ImageMagick
    results = render_pdf_with_magick(pdf, out_d, selected_pages, dpi=dpi, fmt=fmt, prefix=pref)
    if results:
        return results

    raise RuntimeError("Failed to render PDF to images. Ensure `pymupdf` or `magick` is installed.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert PDF to high-resolution images for multimodal models.")
    parser.add_argument("pdf", help="Path to PDF file")
    parser.add_argument("-o", "--output-dir", help="Directory to save images (default: <pdf_name>_images)")
    parser.add_argument("-p", "--pages", help="Pages to render (e.g., '1', '1-3', '1,3,5', or 'all')")
    parser.add_argument("--dpi", type=int, default=150, help="Image resolution DPI (default: 150)")
    parser.add_argument("--format", choices=["png", "jpg"], default="png", help="Image format (default: png)")
    parser.add_argument("--prefix", help="Image filename prefix")

    args = parser.parse_args()
    try:
        images = pdf_to_images(
            args.pdf,
            output_dir=args.output_dir,
            pages=args.pages,
            dpi=args.dpi,
            fmt=args.format,
            prefix=args.prefix,
        )
        print(f"SUCCESS: Rendered {len(images)} images to: {args.output_dir or (Path(args.pdf).stem + '_images')}")
        for img in images:
            print(f"  - {img}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
