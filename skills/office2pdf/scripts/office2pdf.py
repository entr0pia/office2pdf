"""
office2pdf.py
Unified Command Line Interface for Office document structural inspection,
PDF conversion, and multimodal image rendering.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

# Support running both as package and as standalone script
try:
    from .inspect_office import inspect_file, format_inspection_output
    from .convert_office import convert_to_pdf
    from .pdf_to_images import pdf_to_images
except ImportError:
    from inspect_office import inspect_file, format_inspection_output
    from convert_office import convert_to_pdf
    from pdf_to_images import pdf_to_images


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="office2pdf",
        description="Fast structural inspection and high-fidelity PDF/image conversion for Word, Excel, and PowerPoint documents.",
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Command: inspect
    p_inspect = subparsers.add_parser("inspect", help="Quickly inspect document structure without full visual rendering.")
    p_inspect.add_argument("file", help="Office file path (.docx, .xlsx, .pptx, etc.)")
    p_inspect.add_argument("--json", action="store_true", help="Output JSON format")
    p_inspect.add_argument("--rows", type=int, default=5, help="Preview rows for Excel sheets (default: 5)")

    # Command: pdf
    p_pdf = subparsers.add_parser("pdf", help="Convert Word, Excel, or PowerPoint to PDF.")
    p_pdf.add_argument("file", help="Office file path")
    p_pdf.add_argument("-o", "--output", help="Output PDF file path")
    p_pdf.add_argument("--sheet", help="For Excel: convert only the specified sheet name")
    p_pdf.add_argument(
        "--engine",
        choices=["auto", "com", "browser", "soffice"],
        default="auto",
        help="Engine to use: 'com' (native Office), 'browser' (headless Edge/Chrome), 'soffice' (LibreOffice), 'auto' (default)",
    )

    # Command: image
    p_image = subparsers.add_parser("image", help="Convert Office file or PDF directly to images (PNG/JPG).")
    p_image.add_argument("file", help="Office or PDF file path")
    p_image.add_argument("-o", "--output-dir", help="Output directory for images")
    p_image.add_argument("-p", "--pages", help="Page range to export (e.g., '1', '1-3', '1,3,5', or 'all')")
    p_image.add_argument("--dpi", type=int, default=150, help="Image resolution DPI (default: 150)")
    p_image.add_argument("--format", choices=["png", "jpg"], default="png", help="Image format (default: png)")
    p_image.add_argument("--sheet", help="For Excel: convert only the specified sheet")
    p_image.add_argument("--keep-pdf", action="store_true", help="Keep intermediate PDF in same directory as office file")

    # Shorthand preprocessing: if first arg is an existing file or option, route intelligently
    argv = sys.argv[1:]
    if argv and argv[0] not in ["inspect", "pdf", "image", "-h", "--help"] and not argv[0].startswith("-"):
        # Default shorthand: treat as `pdf` command
        argv.insert(0, "pdf")

    args = parser.parse_args(argv)

    # Route based on subcommands
    if args.command == "inspect":
        import json
        res = inspect_file(args.file, max_preview_rows=args.rows)
        if args.json:
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            print(format_inspection_output(res))
        return

    if args.command == "pdf":
        out = convert_to_pdf(args.file, output_path=args.output, sheet_name=args.sheet, engine=args.engine)
        print(f"SUCCESS: {out}")
        return

    if args.command == "image":
        import tempfile
        src = Path(args.file).resolve()
        # If source is already a PDF, directly render to images
        if src.suffix.lower() == ".pdf":
            imgs = pdf_to_images(src, output_dir=args.output_dir, pages=args.pages, dpi=args.dpi, fmt=args.format)
        else:
            # Convert office file to intermediate PDF safely
            if getattr(args, "keep_pdf", False):
                intermediate_pdf = src.with_suffix(".pdf")
                pdf_path = convert_to_pdf(src, output_path=intermediate_pdf, sheet_name=args.sheet)
                imgs = pdf_to_images(pdf_path, output_dir=args.output_dir, pages=args.pages, dpi=args.dpi, fmt=args.format, prefix=src.stem)
            else:
                with tempfile.TemporaryDirectory() as td:
                    intermediate_pdf = Path(td) / f"{src.stem}.pdf"
                    pdf_path = convert_to_pdf(src, output_path=intermediate_pdf, sheet_name=args.sheet)
                    imgs = pdf_to_images(pdf_path, output_dir=args.output_dir, pages=args.pages, dpi=args.dpi, fmt=args.format, prefix=src.stem)
        print(f"SUCCESS: Rendered {len(imgs)} images:")
        for img in imgs:
            print(f"  - {img}")
        return
    parser.print_help()



if __name__ == "__main__":
    main()
