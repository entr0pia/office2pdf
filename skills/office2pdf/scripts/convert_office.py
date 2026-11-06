"""
convert_office.py
High-fidelity conversion of Word, Excel, and PowerPoint documents into PDF.
Multi-engine support:
  1. Windows Office COM (Microsoft Word/Excel/PowerPoint) - Highest visual fidelity
  2. Headless Browser (Edge/Chrome) + HTML/Pandoc - Universal fallback without Office COM
  3. LibreOffice (soffice) - Cross-platform CLI fallback
"""

from __future__ import annotations

import argparse
import html as html_lib
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional


# =====================================================================
# Browser and Tool Discovery
# =====================================================================

def find_headless_browser() -> Optional[Path]:
    """Find installed Edge or Chrome executable for headless HTML-to-PDF printing."""
    candidates = [
        os.environ.get("OFFICE2PDF_BROWSER"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        shutil.which("msedge"),
        shutil.which("chrome"),
        shutil.which("chromium"),
        shutil.which("google-chrome"),
    ]
    for c in candidates:
        if c and Path(c).is_file():
            return Path(c)
    return None


def print_html_to_pdf(html_path: Path, output_pdf_path: Path, landscape: bool = False) -> bool:
    """Print an HTML file to PDF using headless Edge or Chrome."""
    browser = find_headless_browser()
    if not browser:
        return False

    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    html_url = html_path.resolve().as_uri()

    cmd = [
        str(browser),
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={output_pdf_path.resolve()}",
        html_url,
    ]

    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        return output_pdf_path.exists() and output_pdf_path.stat().st_size > 0
    except Exception as e:
        print(f"[Fallback Browser] Error running browser print: {e}", file=sys.stderr)
        return False


# =====================================================================
# Engine 1: Windows Office COM
# =====================================================================

def convert_word_com(input_path: Path, output_path: Path) -> bool:
    """Convert Word document (.docx, .doc) to PDF via Word COM automation."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return False

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        doc = word.Documents.Open(str(input_path.resolve()), ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False)
        # 17 corresponds to wdExportFormatPDF
        doc.ExportAsFixedFormat(
            OutputFileName=str(output_path.resolve()),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,  # wdExportOptimizeForPrint
        )
        return output_path.exists() and output_path.stat().st_size > 0
    except Exception as e:
        print(f"[Word COM] Conversion notice: {e}", file=sys.stderr)
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(0)  # wdDoNotSaveChanges
            except Exception:
                pass
            doc = None
        if word is not None:
            try:
                word.Quit(0)  # wdDoNotSaveChanges
            except Exception:
                pass
            word = None
        import gc
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convert_excel_com(
    input_path: Path,
    output_path: Path,
    sheet_name: Optional[str] = None,
    fit_to_page: bool = True,
) -> bool:
    """Convert Excel workbook (.xlsx, .xls) to PDF via Excel COM automation."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return False

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(input_path.resolve()), ReadOnly=True, UpdateLinks=0)

        if fit_to_page:
            sheets_to_format = [wb.Sheets(sheet_name)] if sheet_name else wb.Sheets
            for ws in sheets_to_format:
                try:
                    ws.PageSetup.Orientation = 2  # xlLandscape
                    ws.PageSetup.Zoom = False
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                except Exception:
                    pass

        # 0 corresponds to xlTypePDF
        if sheet_name:
            ws = wb.Sheets(sheet_name)
            ws.ExportAsFixedFormat(0, str(output_path.resolve()))
        else:
            wb.ExportAsFixedFormat(0, str(output_path.resolve()))

        return output_path.exists() and output_path.stat().st_size > 0
    except Exception as e:
        print(f"[Excel COM] Notice (falling back if needed): {e}", file=sys.stderr)
        return False
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
            wb = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
        import gc
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convert_powerpoint_com(input_path: Path, output_path: Path) -> bool:
    """Convert PowerPoint presentation (.pptx, .ppt) to PDF via PowerPoint COM automation."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return False

    pythoncom.CoInitialize()
    ppt = None
    pres = None
    try:
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        pres = ppt.Presentations.Open(str(input_path.resolve()), WithWindow=False, ReadOnly=True)
        # 32 corresponds to ppSaveAsPDF
        pres.SaveAs(str(output_path.resolve()), 32)
        return output_path.exists() and output_path.stat().st_size > 0
    except Exception as e:
        print(f"[PowerPoint COM] Notice: {e}", file=sys.stderr)
        return False
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
            pres = None
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass
            ppt = None
        import gc
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _run_com_isolated(
    kind: str,
    input_path: Path,
    output_path: Path,
    sheet_name: Optional[str] = None,
    timeout: int = 45,
) -> bool:
    """
    Execute COM conversion inside an isolated subprocess.
    This guarantees COM apartment teardown safety, prevents RPC_E_DISCONNECTED (0x80010108)
    exceptions from polluting pytest or host processes, and kills hanging Office dialogs.
    """
    if sys.platform != "win32":
        return False

    cmd = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--internal-com-worker",
        kind,
        str(input_path.resolve()),
        str(output_path.resolve()),
    ]
    if sheet_name:
        cmd.extend(["--sheet", sheet_name])

    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if res.returncode == 0 and output_path.exists() and output_path.stat().st_size > 0:
            return True
        if res.stderr:
            for line in res.stderr.splitlines():
                if "Notice" in line or "Error" in line:
                    print(line, file=sys.stderr)
        return False
    except subprocess.TimeoutExpired:
        print(f"[{kind.capitalize()} COM] Timed out after {timeout}s", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[{kind.capitalize()} COM] Worker invocation notice: {e}, attempting in-process fallback", file=sys.stderr)
        if kind == "word":
            return convert_word_com(input_path, output_path)
        elif kind == "excel":
            return convert_excel_com(input_path, output_path, sheet_name=sheet_name)
        elif kind == "powerpoint":
            return convert_powerpoint_com(input_path, output_path)
        return False


# =====================================================================
# Engine 2: Fallback (Headless Browser / Pandoc / OpenPyXL)
# =====================================================================

def convert_excel_fallback(
    input_path: Path,
    output_path: Path,
    sheet_name: Optional[str] = None,
) -> bool:
    """Convert Excel to styled HTML table preserving merged cells and highlights, then print to PDF."""
    try:
        import openpyxl
    except ImportError:
        return False

    wb = openpyxl.load_workbook(str(input_path.resolve()), data_only=True)
    sheets = [sheet_name] if sheet_name else wb.sheetnames

    html_parts = [
        "<!DOCTYPE html>",
        "<html><head><meta charset='utf-8'>",
        "<style>",
        "  @page { size: A3 landscape; margin: 10mm; }",
        "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'SimSun', sans-serif; font-size: 11px; margin: 0; padding: 12px; }",
        "  .sheet-title { font-size: 16px; font-weight: bold; margin: 16px 0 8px 0; color: #1f2937; }",
        "  table { border-collapse: collapse; width: 100%; table-layout: fixed; margin-bottom: 24px; page-break-after: auto; }",
        "  th, td { border: 1px solid #cbd5e1; padding: 6px 8px; vertical-align: middle; word-wrap: break-word; line-height: 1.4; }",
        "  th { background-color: #f1f5f9; font-weight: bold; text-align: center; color: #0f172a; }",
        "  .cell-yellow { background-color: #fef08a; }",
        "  .cell-green { background-color: #bbf7d0; }",
        "  .cell-blue { background-color: #bfdbfe; }",
        "  .center { text-align: center; }",
        "  .left { text-align: left; }",
        "  .right { text-align: right; }",
        "</style></head><body>",
    ]

    for s_name in sheets:
        if s_name not in wb.sheetnames:
            continue
        ws = wb[s_name]
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        if max_r == 0 or max_c == 0:
            continue

        # Map merged cell spans
        master_cells = {}
        skipped_cells = set()
        for mrange in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = mrange.min_col, mrange.min_row, mrange.max_col, mrange.max_row
            rowspan = max_row - min_row + 1
            colspan = max_col - min_col + 1
            master_cells[(min_row, min_col)] = (rowspan, colspan)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if (r, c) != (min_row, min_col):
                        skipped_cells.add((r, c))

        html_parts.append(f"<div class='sheet-title'>Sheet: {html_lib.escape(s_name)}</div>")
        html_parts.append("<table>")

        for r in range(1, max_r + 1):
            row_vals = [ws.cell(r, c).value for c in range(1, max_c + 1)]
            if not any(v is not None and str(v).strip() != "" for v in row_vals):
                continue

            html_parts.append("<tr>")
            for c in range(1, max_c + 1):
                if (r, c) in skipped_cells:
                    continue

                cell = ws.cell(r, c)
                val = cell.value
                txt = "" if val is None else str(val)
                txt = html_lib.escape(txt).replace("\n", "<br>")

                classes = []
                # Check background fill
                if cell.fill and cell.fill.fgColor:
                    color_val = str(cell.fill.fgColor.value or "").upper()
                    if "FF00" in color_val or "FFFF00" in color_val:
                        classes.append("cell-yellow")
                    elif "00FF00" in color_val or "A9D08E" in color_val:
                        classes.append("cell-green")
                    elif "00B0F0" in color_val or "5B9BD5" in color_val:
                        classes.append("cell-blue")

                tag = "th" if r <= 2 else "td"
                cls_attr = f" class='{' '.join(classes)}'" if classes else ""
                span_attrs = ""
                if (r, c) in master_cells:
                    rowspan, colspan = master_cells[(r, c)]
                    if rowspan > 1:
                        span_attrs += f" rowspan='{rowspan}'"
                    if colspan > 1:
                        span_attrs += f" colspan='{colspan}'"

                html_parts.append(f"<{tag}{cls_attr}{span_attrs}>{txt}</{tag}>")
            html_parts.append("</tr>")

        html_parts.append("</table>")

    html_parts.append("</body></html>")
    full_html = "\n".join(html_parts)

    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
        tf.write(full_html)
        tmp_html_path = Path(tf.name)

    try:
        success = print_html_to_pdf(tmp_html_path, output_path, landscape=True)
        return success
    finally:
        if tmp_html_path.exists():
            tmp_html_path.unlink()


def convert_word_fallback(input_path: Path, output_path: Path) -> bool:
    """Convert Word .docx to HTML via pandoc or python-docx, preserving flow order, then print to PDF."""
    # Try Pandoc first
    pandoc_bin = shutil.which("pandoc")
    if pandoc_bin:
        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
            tmp_html_path = Path(tf.name)
        try:
            cmd = [
                pandoc_bin,
                str(input_path.resolve()),
                "-s",
                "-o", str(tmp_html_path),
                "--metadata", "title=Document",
            ]
            subprocess.run(cmd, check=True, capture_output=True)
            if print_html_to_pdf(tmp_html_path, output_path, landscape=False):
                return True
        except Exception:
            pass
        finally:
            if tmp_html_path.exists():
                tmp_html_path.unlink()

    # Fallback to python-docx HTML generation preserving document order
    try:
        import docx
    except ImportError:
        return False

    doc = docx.Document(str(input_path.resolve()))
    html_parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'>",
        "<style>",
        "  @page { size: A4 portrait; margin: 20mm; }",
        "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'SimSun', serif; font-size: 14px; line-height: 1.6; padding: 20px; }",
        "  h1 { font-size: 22px; color: #1e3a8a; border-bottom: 2px solid #3b82f6; padding-bottom: 6px; }",
        "  h2 { font-size: 18px; color: #1e40af; margin-top: 20px; }",
        "  h3 { font-size: 16px; color: #1d4ed8; }",
        "  table { border-collapse: collapse; width: 100%; margin: 16px 0; }",
        "  th, td { border: 1px solid #94a3b8; padding: 6px 10px; }",
        "  th { background-color: #f1f5f9; font-weight: bold; }",
        "</style></head><body>",
    ]

    for child in doc.element.body:
        if child.tag.endswith("p"):
            p = docx.text.paragraph.Paragraph(child, doc)
            txt = p.text.strip()
            if not txt:
                continue
            style = p.style.name if p.style else ""
            s_lower = style.lower()
            is_h1 = s_lower == "heading 1" or "标题 1" in style or "标题1" in style
            is_h2 = s_lower == "heading 2" or "标题 2" in style or "标题2" in style
            is_h3 = s_lower == "heading 3" or "标题 3" in style or "标题3" in style
            is_title = s_lower in ["title", "subtitle"] or "标题" in style

            run_parts = []
            for run in p.runs:
                r_txt = html_lib.escape(run.text).replace("\n", "<br>")
                if not r_txt:
                    continue
                if run.bold:
                    r_txt = f"<strong>{r_txt}</strong>"
                if run.italic:
                    r_txt = f"<em>{r_txt}</em>"
                if run.underline:
                    r_txt = f"<u>{r_txt}</u>"
                run_parts.append(r_txt)
            rendered_p = "".join(run_parts) if run_parts else html_lib.escape(txt)

            if is_h1:
                html_parts.append(f"<h1>{rendered_p}</h1>")
            elif is_h2:
                html_parts.append(f"<h2>{rendered_p}</h2>")
            elif is_h3:
                html_parts.append(f"<h3>{rendered_p}</h3>")
            elif is_title:
                html_parts.append(f"<h1 style='text-align:center;'>{rendered_p}</h1>")
            else:
                html_parts.append(f"<p>{rendered_p}</p>")

        elif child.tag.endswith("tbl"):
            tbl = docx.table.Table(child, doc)
            html_parts.append("<table>")
            for r_idx, row in enumerate(tbl.rows):
                html_parts.append("<tr>")
                for cell in row.cells:
                    c_txt = html_lib.escape(cell.text.strip()).replace("\n", "<br>")
                    tag = "th" if r_idx == 0 else "td"
                    html_parts.append(f"<{tag}>{c_txt}</{tag}>")
                html_parts.append("</tr>")
            html_parts.append("</table>")

    html_parts.append("</body></html>")
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
        tf.write("\n".join(html_parts))
        tmp_html = Path(tf.name)

    try:
        return print_html_to_pdf(tmp_html, output_path, landscape=False)
    finally:
        if tmp_html.exists():
            tmp_html.unlink()


# =====================================================================
# Engine 3: LibreOffice CLI
# =====================================================================

def convert_via_libreoffice(input_path: Path, output_path: Path) -> bool:
    """Convert via LibreOffice headless CLI (soffice)."""
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        return False

    out_dir = output_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice_bin,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir.resolve()),
        str(input_path.resolve()),
    ]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        expected = out_dir / (input_path.stem + ".pdf")
        if expected.exists():
            if expected.resolve() != output_path.resolve():
                shutil.move(str(expected), str(output_path))
            return True
        return False
    except Exception as e:
        print(f"[LibreOffice] Failed: {e}", file=sys.stderr)
        return False


# =====================================================================
# Unified Converter Dispatcher
# =====================================================================

def convert_to_pdf(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
    sheet_name: Optional[str] = None,
    engine: Optional[str] = None,
) -> Path:
    """
    Convert an Office document (.docx, .doc, .xlsx, .xls, .pptx, .ppt) to PDF.
    
    Args:
        input_path: Path to source document.
        output_path: Target PDF path. Defaults to input path with .pdf extension.
        sheet_name: For Excel, specify a single sheet name to convert.
        engine: Force specific engine ('com', 'browser', 'soffice', 'auto').
    """
    src = Path(input_path).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Source file not found: {src}")

    dst = Path(output_path).resolve() if output_path else src.with_suffix(".pdf")
    dst.parent.mkdir(parents=True, exist_ok=True)

    ext = src.suffix.lower()
    selected_engine = (engine or "auto").lower()

    is_windows = sys.platform == "win32"

    if ext in [".docx", ".doc"]:
        if selected_engine in ["auto", "com"] and is_windows:
            if _run_com_isolated("word", src, dst):
                return dst
        if selected_engine in ["auto", "soffice"]:
            if convert_via_libreoffice(src, dst):
                return dst
        if selected_engine in ["auto", "browser"]:
            if convert_word_fallback(src, dst):
                return dst

    elif ext in [".xlsx", ".xls", ".xlsm"]:
        if selected_engine in ["auto", "com"] and is_windows:
            if _run_com_isolated("excel", src, dst, sheet_name=sheet_name):
                return dst
        if selected_engine in ["auto", "soffice"]:
            if convert_via_libreoffice(src, dst):
                return dst
        if selected_engine in ["auto", "browser"]:
            if convert_excel_fallback(src, dst, sheet_name=sheet_name):
                return dst

    elif ext in [".pptx", ".ppt"]:
        if selected_engine in ["auto", "com"] and is_windows:
            if _run_com_isolated("powerpoint", src, dst):
                return dst
        if selected_engine in ["auto", "soffice"]:
            if convert_via_libreoffice(src, dst):
                return dst

    raise RuntimeError(
        f"Failed to convert {src.name} to PDF using available engines "
        f"(tried COM={is_windows}, LibreOffice={bool(shutil.which('soffice'))}, Browser={bool(find_headless_browser())})."
    )


def main() -> None:
    if "--internal-com-worker" in sys.argv:
        p_worker = argparse.ArgumentParser()
        p_worker.add_argument("--internal-com-worker", choices=["word", "excel", "powerpoint"], required=True)
        p_worker.add_argument("input")
        p_worker.add_argument("output")
        p_worker.add_argument("--sheet", default=None)
        args_w, _ = p_worker.parse_known_args(sys.argv[1:])
        src_p = Path(args_w.input).resolve()
        dst_p = Path(args_w.output).resolve()
        ok = False
        if args_w.internal_com_worker == "word":
            ok = convert_word_com(src_p, dst_p)
        elif args_w.internal_com_worker == "excel":
            ok = convert_excel_com(src_p, dst_p, sheet_name=args_w.sheet)
        elif args_w.internal_com_worker == "powerpoint":
            ok = convert_powerpoint_com(src_p, dst_p)
        sys.exit(0 if ok else 1)

    parser = argparse.ArgumentParser(description="Convert Word, Excel, or PowerPoint files to PDF.")
    parser.add_argument("file", help="Input Office file (.docx, .doc, .xlsx, .xls, .pptx, .ppt)")
    parser.add_argument("-o", "--output", help="Output PDF file path (default: <filename>.pdf)")
    parser.add_argument("--sheet", help="For Excel files: convert only the specified sheet")
    parser.add_argument(
        "--engine",
        choices=["auto", "com", "browser", "soffice"],
        default="auto",
        help="Conversion engine to use (default: auto)",
    )

    args = parser.parse_args()
    try:
        out = convert_to_pdf(args.file, output_path=args.output, sheet_name=args.sheet, engine=args.engine)
        print(f"SUCCESS: Converted to PDF -> {out}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
