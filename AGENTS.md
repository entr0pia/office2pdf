# AGENTS.md

Technical context, architectural design, and actionable guidelines for AI coding agents working on the `office2pdf` repository.

---

## 1. Project Overview

`office2pdf` is a specialized Python toolkit and agent skill designed for multimodal LLMs (e.g., Gemini, Claude, GPT-4o). It enables:
1. **Zero-Token Structural Inspection**: Code-based fast metadata extraction (sheet names, dimensions, heading trees, slide outlines) without burning visual tokens on massive documents.
2. **Multi-Engine PDF Conversion**: High-fidelity conversion of Word (`.docx`, `.doc`), Excel (`.xlsx`, `.xls`), and PowerPoint (`.pptx`, `.ppt`) documents into PDF.
   - Primary: Microsoft Office COM Automation (`pywin32`) on Windows for 100% native layout fidelity.
   - Secondary: Headless Edge/Chrome browser + styled HTML table (`openpyxl`) / Pandoc fallback for environments without Office COM or without installed printer drivers.
   - Tertiary: LibreOffice CLI (`soffice`) for headless Linux/cross-platform environments.
3. **Multimodal Image Rasterization**: High-resolution rendering (150–300 DPI) to PNG/JPG via PyMuPDF (`pymupdf`) with ImageMagick (`magick`) fallback.

---

## 2. Environment & Setup

- **Platform**: Windows 10/11, PowerShell (`pwsh`).
- **Python Manager**: Always use `uv` for virtual environment and dependency management.
- **Character Encoding**: Set `$env:PYTHONUTF8=1` in PowerShell before running Python scripts to prevent encoding issues with non-ASCII filenames or content.

### Setup Commands

```powershell
# Install project dependencies into .venv
uv sync

# If running isolated commands with dependencies
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/
```

### Core Dependencies
- `pywin32`: Windows COM automation for Word, Excel, PowerPoint.
- `openpyxl`: Excel inspection and HTML table generation for headless browser fallback.
- `python-docx`: Word inspection and structured HTML generation.
- `python-pptx`: PowerPoint inspection and slide outline extraction.
- `pymupdf`: High-speed PDF-to-image rasterization.

---

## 3. Development Workflow & CLI Usage

The repository structure follows a clean layout:
```
office2pdf/
├── skills/
│   └── office2pdf/
│       ├── SKILL.md                # Agent skill specification
│       └── scripts/
│           ├── __init__.py
│           ├── office2pdf.py       # Main CLI dispatcher (inspect, pdf, image)
│           ├── inspect_office.py   # Structural inspector for Word, Excel, PPT
│           ├── convert_office.py   # Multi-engine PDF conversion pipeline
│           └── pdf_to_images.py    # PyMuPDF / ImageMagick image rasterizer
├── tests/
│   ├── generate_samples.py # Generates mock .docx, .xlsx, .pptx files
│   ├── test_inspect.py     # Structural inspection tests
│   ├── test_convert.py     # Conversion pipeline tests
│   ├── test_pdf_to_images.py # Image rasterization tests
│   └── test_cli.py         # Subcommand & CLI tests
├── pyproject.toml          # Hatchling package specification
├── README.md               # Human-facing overview
└── AGENTS.md               # Agent guide (this file)
```

### CLI Commands

```powershell
# 1. Structural inspection (human-readable summary)
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf inspect path/to/document.xlsx

# 2. Structural inspection (JSON format)
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf inspect path/to/document.docx --json

# 3. Convert Office document to PDF
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf pdf path/to/document.docx -o output.pdf

# 4. Convert specific Excel sheet to PDF
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf pdf path/to/document.xlsx --sheet "Sheet1"

# 5. Render directly to high-res images (PNG/JPG)
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf image path/to/presentation.pptx --dpi 200

# 6. Render specific PDF pages
$env:PYTHONUTF8=1; uv run python -m scripts.office2pdf image path/to/document.pdf --pages 1-3 --dpi 150
```

---

## 4. Testing Instructions

All automated tests use `pytest` and programmatic document generation to ensure zero external file dependencies.

### Running Tests

```powershell
# Generate test sample documents in test_data/
$env:PYTHONUTF8=1; uv run python tests/generate_samples.py

# Run entire test suite
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/ -v

# Run targeted test suites
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/test_inspect.py -v
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/test_convert.py -v
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/test_pdf_to_images.py -v
$env:PYTHONUTF8=1; uv run --with pytest pytest tests/test_cli.py -v
```

---

## 5. Architectural Conventions & Critical Gotchas

### Windows COM Automation Rules
1. **Thread Apartment Initialization**: Always pair `pythoncom.CoInitialize()` and `pythoncom.CoUninitialize()` around COM dispatches.
2. **Explicit Teardown & Garbage Collection**:
   - COM objects (Document, Workbook, Presentation) must be explicitly closed (`.Close(False)`).
   - Variables must be cleared (`del doc; doc = None`) followed by `gc.collect()` before calling `app.Quit()`.
   - Failing to explicitly clear references before quitting out-of-process servers (`WINWORD.EXE`, `EXCEL.EXE`, `POWERPNT.EXE`) causes `RPC_E_DISCONNECTED` (`0x80010108`) or dangling background processes.
3. **Excel COM Missing Printer Driver**:
   - On Windows machines where no physical or virtual printer driver is installed, Excel COM's `ExportAsFixedFormat` will fail with an ActiveX exception (`尚未安装打印机`).
   - `scripts/convert_office.py` catches this exception and seamlessly falls back to `convert_excel_fallback()` (rendering an HTML table via `openpyxl` with preserved cell highlights and printing to PDF via headless Microsoft Edge). Always preserve this fallback chain.

### Multimodal AI Workflow Best Practice
- **Inspect Before Rendering**: When analyzing unfamiliar Office documents, never jump straight into converting a multi-hundred-page or multi-sheet file. Run `inspect` first to understand the sheet list, heading outline, and page counts, then target specific sheets or pages.
