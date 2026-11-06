# office2pdf

A specialized toolkit and Agent Skill tailored for Multimodal LLMs (Gemini, Claude, GPT-4o) to visually inspect, structure-parse, and convert Microsoft Office documents (Word, Excel, PowerPoint) into high-fidelity PDFs and crisp images.

---

## What It Can Do

1. **Zero-Token Structural Inspection (Fast Code Parsing)**
   - **Excel (`.xlsx`, `.xls`, `.xlsm`)**: Instantly inspect worksheet names, active data dimensions, and sample the first N header rows without converting or burning visual tokens.
   - **Word (`.docx`, `.doc`)**: Extract the full heading hierarchy (`Title`, `Heading 1-3`, Chinese `标题 1-3`) into a Table of Contents tree and inspect table dimensions.
   - **PowerPoint (`.pptx`, `.ppt`)**: Scan slide counts, slide titles, shapes/tables/images counts, and speaker notes.

2. **Multi-Engine High-Fidelity PDF Conversion**
   - **Microsoft Office COM (Windows)**: Uses native `WINWORD.EXE`, `EXCEL.EXE`, and `POWERPNT.EXE` for 100% pixel-perfect layout fidelity.
   - **Headless Browser Fallback**: When Office or printer drivers are absent, converts Excel sheets into semantic styled HTML tables (retaining cell highlights, borders, and merges) and prints them via headless Microsoft Edge or Google Chrome.
   - **LibreOffice CLI (`soffice`)**: Fallback for Linux or headless environments.

3. **Multimodal Image Rasterization**
   - Renders Office documents or existing PDFs directly into high-resolution PNG or JPEG images (150–300 DPI) using PyMuPDF (`pymupdf`) with ImageMagick (`magick`) fallback.
   - Supports page range selection (e.g. `--pages 1,3-5`) to minimize token consumption.

4. **Robust COM Process Isolation**
   - Runs Windows COM operations in isolated watchdog child processes with automatic garbage collection, eliminating lingering background processes (`RPC_E_DISCONNECTED` / `0x80010108` errors).

---

## Prerequisites

- **Operating System**: Windows 10/11 (recommended for native Office COM fidelity), Linux, or macOS.
- **Python**: Python 3.10+ (managed via [uv](https://docs.astral.sh/uv/) recommended).
- **Software Dependencies**:
  - *Recommended on Windows*: Microsoft Office (Word, Excel, PowerPoint) installed.
  - *Browser Fallback*: Microsoft Edge or Google Chrome (for headless HTML-to-PDF printing).

---

## Installation

### Option 1: Install as an Agent Skill (Recommended for AI Agents)

Install into your AI coding assistant (Antigravity, Claude Code, Cursor) using the `skills` CLI:

```bash
# Global installation for all AI agents
npx skills add entr0pia/office2pdf -g

# Or workspace-level installation
npx skills add entr0pia/office2pdf
```

Alternatively, copy the `skills/office2pdf` folder directly into your global agent skills directory:
- **Windows**: `~/.agents/skills/office2pdf`
- **Linux/macOS**: `~/.agents/skills/office2pdf`

### Option 2: Local CLI / Developer Setup

Clone the repository and install dependencies using `uv`:

```bash
# Clone the repository
git clone https://github.com/entr0pia/office2pdf.git
cd office2pdf

# Sync dependencies in a virtual environment
uv sync
```

---

## Quick Usage (CLI)

Thanks to [PEP 723](https://peps.python.org/pep-0723/) inline script metadata, `uv run` **automatically resolves and caches all required dependencies** in a clean, isolated environment on first execution. You can run the script directly from any working directory without pre-activating a virtual environment:

### 1. Structural Inspection (Zero Vision Tokens)
```bash
# Human-readable summary
uv run skills/office2pdf/scripts/office2pdf.py inspect sample.xlsx

# JSON output for automated scripting
uv run skills/office2pdf/scripts/office2pdf.py inspect sample.docx --json

# Inspect Excel with first 5 rows preview
uv run skills/office2pdf/scripts/office2pdf.py inspect financial_report.xlsx --rows 5
```

### 2. High-Fidelity PDF Conversion
```bash
# Automatic engine selection (Word, Excel, or PPT)
uv run skills/office2pdf/scripts/office2pdf.py pdf report.docx -o report.pdf

# Convert a specific Excel worksheet
uv run skills/office2pdf/scripts/office2pdf.py pdf report.xlsx --sheet "Q3 Financials" -o q3.pdf
```

### 3. Direct Image Rendering (For Vision Models)
```bash
# Convert entire presentation to 200 DPI PNG images
uv run skills/office2pdf/scripts/office2pdf.py image presentation.pptx --dpi 200 -o ./output_images

# Render specific pages from a document or PDF
uv run skills/office2pdf/scripts/office2pdf.py image document.pdf --pages 1-3 --dpi 150
```

---

## Project Structure

```
office2pdf/
├── skills/
│   └── office2pdf/
│       ├── SKILL.md                # Agent Skill prompt & operational instructions
│       └── scripts/                # Preset python scripts
│           ├── __init__.py
│           ├── office2pdf.py       # Main CLI dispatcher (inspect, pdf, image)
│           ├── inspect_office.py   # Zero-token structural inspector
│           ├── convert_office.py   # Multi-engine PDF conversion pipeline
│           └── pdf_to_images.py    # PyMuPDF / ImageMagick rasterizer
├── tests/                          # Comprehensive pytest test suite (22 tests)
├── test_data/                      # Mock document generator & fixtures
├── pyproject.toml                  # Python package and pytest configuration
├── README.md                       # Human-facing documentation (this file)
└── AGENTS.md                       # Developer & AI Agent architectural guidelines
```

---

## License

MIT License.
