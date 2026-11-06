# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1.2",
#     "python-docx>=1.1.0",
#     "python-pptx>=0.6.23",
# ]
# ///

"""
inspect_office.py
Fast structural inspection for Office documents (Word, Excel, PowerPoint)
without full visual rendering or token-heavy conversion.
Supports modern OpenXML (.docx, .xlsx, .pptx) and legacy formats (.doc, .xls, .ppt).
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Ensure script directory and parent directory are on sys.path
_this_dir = Path(__file__).resolve().parent
for _p in [str(_this_dir), str(_this_dir.parent)]:
    if _p not in sys.path:
        sys.path.insert(0, _p)


def inspect_excel(file_path: Path, max_preview_rows: int = 5) -> Dict[str, Any]:
    """Inspect Excel spreadsheet structure (.xlsx, .xlsm)."""
    import openpyxl

    wb = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True)
    sheets_info: List[Dict[str, Any]] = []

    for name in wb.sheetnames:
        sheet = wb[name]
        max_r = sheet.max_row or 0
        max_c = sheet.max_column or 0

        # Read preview rows safely even if dimension ref is missing
        preview_rows: List[List[str]] = []
        try:
            for r_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                if r_idx > max_preview_rows:
                    break
                cols_limit = min(max_c, 20) if max_c > 0 else 20
                row_vals = ["" if v is None else str(v).strip() for v in row[:cols_limit]]
                if any(row_vals):
                    preview_rows.append(row_vals)
                    if max_c == 0:
                        max_c = max(max_c, len(row_vals))
            if max_r == 0 and preview_rows:
                max_r = len(preview_rows)
        except Exception:
            pass

        sheets_info.append({
            "name": name,
            "max_rows": max_r,
            "max_columns": max_c,
            "preview_rows": preview_rows,
        })

    wb.close()
    return {
        "file_type": "excel",
        "file_path": str(file_path),
        "file_name": file_path.name,
        "sheet_count": len(sheets_info),
        "sheets": sheets_info,
    }


def inspect_word(file_path: Path) -> Dict[str, Any]:
    """Inspect Word document structure (.docx)."""
    import docx

    doc = docx.Document(str(file_path))
    headings: List[Dict[str, Any]] = []
    paragraphs_count = len(doc.paragraphs)
    tables_count = len(doc.tables)

    for p_idx, p in enumerate(doc.paragraphs):
        style_name = p.style.name if p.style else ""
        text = p.text.strip()
        if not text:
            continue
        s_lower = style_name.lower()
        if (
            s_lower.startswith("heading")
            or s_lower in ["title", "subtitle"]
            or "标题" in style_name
            or "副标题" in style_name
        ):
            headings.append({
                "index": p_idx,
                "style": style_name,
                "text": text,
            })

    tables_info: List[Dict[str, Any]] = []
    for t_idx, tbl in enumerate(doc.tables):
        r_cnt = len(tbl.rows)
        try:
            c_cnt = len(tbl.columns) if tbl.rows else 0
        except Exception:
            c_cnt = max((len(r.cells) for r in tbl.rows), default=0)
        first_row = []
        if tbl.rows:
            first_row = [cell.text.strip() for cell in tbl.rows[0].cells[:10]]
        tables_info.append({
            "index": t_idx + 1,
            "rows": r_cnt,
            "columns": c_cnt,
            "headers": first_row,
        })

    # Metadata
    core_props = doc.core_properties
    meta = {
        "title": core_props.title or "",
        "author": core_props.author or "",
        "created": str(core_props.created or ""),
        "modified": str(core_props.modified or ""),
    }

    return {
        "file_type": "word",
        "file_path": str(file_path),
        "file_name": file_path.name,
        "paragraph_count": paragraphs_count,
        "heading_count": len(headings),
        "headings": headings,
        "table_count": tables_count,
        "tables": tables_info,
        "metadata": meta,
    }


def inspect_powerpoint(file_path: Path) -> Dict[str, Any]:
    """Inspect PowerPoint presentation structure (.pptx)."""
    import pptx

    prs = pptx.Presentation(str(file_path))
    slides_info: List[Dict[str, Any]] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        title = ""
        shape_texts: List[str] = []
        tables_count = 0
        images_count = 0

        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    if not title and (shape == slide.shapes.title or "title" in shape.name.lower()):
                        title = txt.splitlines()[0]
                    else:
                        shape_texts.append(txt)
            if shape.has_table:
                tables_count += 1
            if shape.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                images_count += 1

        # Fallback for title if not set
        if not title and shape_texts:
            title = shape_texts[0].splitlines()[0]

        # Speaker notes
        notes_text = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()

        slides_info.append({
            "slide_number": s_idx,
            "title": title or "(Untitled)",
            "shape_count": len(slide.shapes),
            "table_count": tables_count,
            "image_count": images_count,
            "text_snippets": [t[:100] for t in shape_texts[:3]],
            "notes": notes_text[:200] if notes_text else "",
        })

    return {
        "file_type": "powerpoint",
        "file_path": str(file_path),
        "file_name": file_path.name,
        "slide_count": len(slides_info),
        "slides": slides_info,
    }


def inspect_legacy_office(file_path: Path) -> Dict[str, Any]:
    """
    Inspect legacy binary Office formats (.doc, .xls, .ppt).
    Provides instant file metadata and conversion guidance without heavy COM overhead.
    """
    ext = file_path.suffix.lower()
    stat = file_path.stat()
    base_info: Dict[str, Any] = {
        "file_type": "legacy_office",
        "file_path": str(file_path),
        "file_name": file_path.name,
        "extension": ext,
        "size_bytes": stat.st_size,
        "message": (
            f"Legacy binary Office format ({ext}). "
            f"For visual analysis, convert to PDF using: uv run python -m scripts.office2pdf pdf \"{file_path.name}\""
        ),
    }

    # If it's .xls, attempt lightweight xlrd inspection if installed
    if ext == ".xls":
        try:
            import xlrd
            book = xlrd.open_workbook(str(file_path.resolve()), on_demand=True)
            sheets_info = [{"name": s} for s in book.sheet_names()]
            base_info["sheet_count"] = len(sheets_info)
            base_info["sheets"] = sheets_info
            base_info["message"] = f"Legacy Excel format (.xls) containing {len(sheets_info)} sheets."
        except Exception:
            pass

    return base_info


def format_inspection_output(data: Dict[str, Any]) -> str:
    """Format inspection result as a clean, human-readable summary."""
    ftype = data.get("file_type")
    lines: List[str] = []
    lines.append(f"=== Document Inspection: {data.get('file_name')} ({ftype.upper()}) ===")
    lines.append(f"Path: {data.get('file_path')}")

    if ftype == "excel":
        lines.append(f"Total Sheets: {data.get('sheet_count')}")
        for s in data.get("sheets", []):
            lines.append(f"\nSheet: [{s['name']}] (Rows: {s['max_rows']}, Columns: {s['max_columns']})")
            if s.get("preview_rows"):
                lines.append("  Header / Preview Rows:")
                for r_idx, r in enumerate(s["preview_rows"], 1):
                    row_str = " | ".join(r[:8])
                    if len(r) > 8:
                        row_str += f" | ... (+{len(r)-8} more cols)"
                    lines.append(f"    R{r_idx}: {row_str}")

    elif ftype == "word":
        lines.append(f"Paragraphs: {data.get('paragraph_count')}, Tables: {data.get('table_count')}")
        meta = data.get("metadata", {})
        if meta.get("title") or meta.get("author"):
            lines.append(f"Title: {meta.get('title', 'N/A')}, Author: {meta.get('author', 'N/A')}")
        lines.append(f"\nHeading Hierarchy ({data.get('heading_count')} headings):")
        if data.get("headings"):
            for h in data["headings"]:
                indent = "  "
                if "2" in h["style"]:
                    indent = "    "
                elif "3" in h["style"]:
                    indent = "      "
                lines.append(f"{indent}* [{h['style']}] {h['text']}")
        else:
            lines.append("  (No structured headings detected)")

        if data.get("tables"):
            lines.append(f"\nTables Outline ({data.get('table_count')} tables):")
            for t in data["tables"]:
                hdr = " | ".join(t["headers"][:6])
                lines.append(f"  Table {t['index']}: {t['rows']} rows x {t['columns']} cols -> [{hdr}]")

    elif ftype == "powerpoint":
        lines.append(f"Total Slides: {data.get('slide_count')}")
        for s in data.get("slides", []):
            lines.append(f"  Slide {s['slide_number']}: \"{s['title']}\" (Shapes: {s['shape_count']}, Tables: {s['table_count']}, Images: {s['image_count']})")
            if s.get("notes"):
                lines.append(f"    Notes: {s['notes']}")

    elif ftype == "legacy_office":
        lines.append(f"Format: Legacy Binary Office ({data.get('extension')})")
        if data.get("message"):
            lines.append(f"Notice: {data.get('message')}")
        if data.get("sheet_count") is not None:
            lines.append(f"Total Sheets: {data.get('sheet_count')}")
            for s in data.get("sheets", []):
                lines.append(f"  Sheet: [{s.get('name')}]")
        if data.get("paragraph_count") is not None:
            lines.append(f"Paragraphs: {data.get('paragraph_count')}, Tables: {data.get('table_count')}")
        if data.get("slide_count") is not None:
            lines.append(f"Total Slides: {data.get('slide_count')}")

    return "\n".join(lines)


def inspect_file(file_path: str | Path, max_preview_rows: int = 5) -> Dict[str, Any]:
    """Inspect any supported office file."""
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    ext = path.suffix.lower()
    if ext in [".xlsx", ".xlsm"]:
        return inspect_excel(path, max_preview_rows=max_preview_rows)
    elif ext in [".docx"]:
        return inspect_word(path)
    elif ext in [".pptx"]:
        return inspect_powerpoint(path)
    elif ext in [".xls", ".doc", ".ppt"]:
        return inspect_legacy_office(path)
    else:
        raise ValueError(f"Unsupported file format: {ext}. Supported formats: .docx, .xlsx, .pptx (and legacy .doc, .xls, .ppt)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect Office document structure (Word, Excel, PowerPoint).")
    parser.add_argument("file", help="Path to Office file (.docx, .xlsx, .pptx, .doc, .xls, .ppt)")
    parser.add_argument("--json", action="store_true", help="Output raw JSON instead of text summary")
    parser.add_argument("--rows", type=int, default=5, help="Number of preview rows for Excel (default: 5)")

    args = parser.parse_args()
    try:
        result = inspect_file(args.file, max_preview_rows=args.rows)
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(format_inspection_output(result))
    except Exception as e:
        print(f"Error inspecting {args.file}: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
