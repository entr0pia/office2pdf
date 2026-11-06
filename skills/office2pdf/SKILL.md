---
name: office2pdf
description: "Convert Word (.docx/.doc), Excel (.xlsx/.xls), and PowerPoint (.pptx/.ppt) documents into PDF or high-resolution images (PNG/JPG) for multimodal AI inspection, combined with fast code-based structural inspection (sheet names, TOC/headings, slide outlines). Make sure to use this skill whenever dealing with Office files to view visual layout, inspect sheets, extract tables, convert slides, or when multimodal vision (Gemini native PDF reading or vision models) is needed instead of raw text dumping."
---

# office2pdf: Multimodal Office Document Pipeline

A comprehensive skill and toolset for multimodal LLMs to inspect, convert, and visually analyze Microsoft Office documents (Word, Excel, PowerPoint).

## Core Philosophy: Inspect First, Then Render

Multimodal large language models (especially Gemini with native PDF support, and vision-enabled models like Claude or GPT-4o) understand complex documents (tables, merged cells, financial sheets, slide diagrams, multi-column reports) **significantly faster and more accurately** when "looking" at rendered PDFs or images compared to parsing messy XML or dumping flat text.

However, **do not blindly convert massive files**. Follow this two-step pattern:
1. **Inspect First (Zero Token Overhead)**: Run code-based inspection to get sheets, table outlines, headings, or slide counts.
2. **Targeted Render**:
   - **Gemini**: Convert target document or sheet to PDF (Gemini natively inspects PDFs).
   - **Vision Models / Grounding**: Convert specific pages or sheets to PNG/JPG images.

---

## Quick Reference

> [!TIP]
> The scripts include [PEP 723](https://peps.python.org/pep-0723/) inline metadata. `uv run` will **automatically install and resolve all required dependencies** in an isolated sandbox on first run—no manual `pip install` or `--with` flags required.
>
> You can run via script path from anywhere:
> `uv run <skill_dir>/scripts/office2pdf.py <command> ...`

| Task | Command |
|---|---|
| **Inspect any Office file** | `uv run <scripts>/office2pdf.py inspect <file>` |
| **Inspect Excel sheets & headers** | `uv run <scripts>/office2pdf.py inspect <file.xlsx> --rows 5` |
| **Inspect as JSON** | `uv run <scripts>/office2pdf.py inspect <file> --json` |
| **Convert to PDF (All formats)** | `uv run <scripts>/office2pdf.py pdf <file> -o <output.pdf>` |
| **Convert specific Excel sheet** | `uv run <scripts>/office2pdf.py pdf <file.xlsx> --sheet "Sheet1"` |
| **Convert directly to Images** | `uv run <scripts>/office2pdf.py image <file> --dpi 200` |
| **Convert specific PDF/doc pages** | `uv run <scripts>/office2pdf.py image <file.pdf> --pages 1,3-5` |

---

## 1. Fast Structural Inspection (Code Parsing)

Use `scripts.inspect_office` to instantly inspect basic structural metadata without token-heavy rendering:

### Excel (.xlsx, .xlsm, .xls)
- Lists all worksheets, row/column dimensions, and non-empty bounds.
- Previews the first N rows of each worksheet (headers, IDs, column names).
```bash
uv run <scripts>/office2pdf.py inspect report.xlsx
```

### Word (.docx, .doc)
- Extracts the heading hierarchy (`Title`, `Heading 1`, `Heading 2`, `Heading 3`) into a clean Table of Contents tree.
- Lists all tables with row count, column count, and sample header cells.
```bash
uv run <scripts>/office2pdf.py inspect contract.docx
```

### PowerPoint (.pptx, .ppt)
- Lists slide count, slide titles, shapes/tables/images count per slide.
- Extracts speaker notes and slide text outlines.
```bash
uv run <scripts>/office2pdf.py inspect presentation.pptx
```

---

## 2. Converting to PDF

For multimodal models that support native PDF inspection (e.g. Gemini 1.5 Pro / Flash, Gemini 2.0), viewing the converted PDF via `read_file` or multimodal tools provides **100% layout fidelity** with zero OCR loss.

### Engine Hierarchy & Automatic Fallback
The conversion script uses an intelligent multi-engine pipeline:
1. **Windows Office COM (`pywin32`)**: Uses installed Microsoft Word, Excel, and PowerPoint for pixel-perfect native output.
2. **Headless Browser & HTML Fallback (`msedge` / `chrome`)**:
   - For Excel: Parses worksheets with `openpyxl`, builds a clean semantic HTML table preserving cell background highlights (e.g. yellow status cells), font weights, borders, and margins, then prints to PDF via headless browser.
   - For Word: Converts docx via Pandoc or python-docx to HTML, then prints to PDF via headless browser.
3. **LibreOffice CLI (`soffice`)**: Cross-platform headless converter when LibreOffice is installed.

### Examples

```bash
# Automatic engine selection (Word, Excel, or PPT)
uv run <scripts>/office2pdf.py pdf document.docx
uv run <scripts>/office2pdf.py pdf spreadsheet.xlsx
uv run <scripts>/office2pdf.py pdf slides.pptx

# Export a single sheet from Excel
uv run <scripts>/office2pdf.py pdf spreadsheet.xlsx --sheet "Q3 Financials" -o q3.pdf

# Force specific engine
uv run <scripts>/office2pdf.py pdf spreadsheet.xlsx --engine browser
```

---

## 3. Converting to High-Resolution Images

For multimodal models that accept images (PNG/JPEG) or when zooming into a specific page/chart:

### Examples

```bash
# Convert all pages/slides to PNG at 150 DPI (default)
uv run <scripts>/office2pdf.py image presentation.pptx

# Convert specific pages at high resolution (200 DPI)
uv run <scripts>/office2pdf.py image report.pdf --pages 1-3 --dpi 200

# Convert directly from Office document with custom output directory
uv run <scripts>/office2pdf.py image contract.docx -o ./contract_pages --format jpg
```

### Rendering Engine
- **PyMuPDF (`pymupdf`)**: Ultra-fast vector-quality rasterization directly to PNG/JPG.
- **ImageMagick (`magick`)**: Automatic fallback when PyMuPDF is unavailable.

---

## 4. Best Practices for Multimodal Agents

1. **Never guess merged cells or table alignments**:
   - If an Excel sheet has complex headers or merged cells, convert the sheet to PDF/image and visually inspect it.
2. **Examine colors and highlights**:
   - Status indicators (red/yellow/green cells) in spreadsheets often carry critical business meaning that flat text dumping loses.
3. **Keep temporary artifacts organized**:
   - Use `-o <path>` to save PDF/images in appropriate workspaces. Clean up scratch files when analysis is complete.
