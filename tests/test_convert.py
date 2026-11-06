"""
Unit tests for convert_office.py.
"""

from pathlib import Path
import pytest
import pymupdf

from scripts.convert_office import convert_to_pdf, convert_excel_fallback, convert_word_fallback
from tests.generate_samples import (
    generate_sample_word,
    generate_sample_excel,
    generate_sample_powerpoint,
)


@pytest.fixture(scope="session")
def test_data_dir(tmp_path_factory) -> Path:
    d = tmp_path_factory.mktemp("convert_data")
    generate_sample_word(d / "test.docx")
    generate_sample_excel(d / "test.xlsx")
    generate_sample_powerpoint(d / "test.pptx")
    return d


def test_convert_word_to_pdf(test_data_dir: Path, tmp_path: Path):
    doc_path = test_data_dir / "test.docx"
    out_pdf = tmp_path / "word.pdf"
    res_pdf = convert_to_pdf(doc_path, output_path=out_pdf)
    assert res_pdf.exists()
    assert res_pdf.stat().st_size > 0

    # Verify PDF validity via PyMuPDF
    with pymupdf.open(str(res_pdf)) as pdf:
        assert len(pdf) >= 1
        page_text = pdf[0].get_text()
        assert "Project Overview Report" in page_text or "Executive Summary" in page_text


def test_convert_excel_to_pdf_auto(test_data_dir: Path, tmp_path: Path):
    excel_path = test_data_dir / "test.xlsx"
    out_pdf = tmp_path / "excel_auto.pdf"
    res_pdf = convert_to_pdf(excel_path, output_path=out_pdf)
    assert res_pdf.exists()
    assert res_pdf.stat().st_size > 0

    with pymupdf.open(str(res_pdf)) as pdf:
        assert len(pdf) >= 1
        page_text = pdf[0].get_text()
        assert "Metrics" in page_text or "Substation Inspection" in page_text


def test_convert_excel_fallback_html(test_data_dir: Path, tmp_path: Path):
    excel_path = test_data_dir / "test.xlsx"
    out_pdf = tmp_path / "excel_fallback.pdf"
    success = convert_excel_fallback(excel_path, out_pdf)
    assert success
    assert out_pdf.exists()
    with pymupdf.open(str(out_pdf)) as pdf:
        assert len(pdf) >= 1
        page_text = pdf[0].get_text()
        assert "Metrics" in page_text or "Scenario Name" in page_text


def test_convert_powerpoint_to_pdf(test_data_dir: Path, tmp_path: Path):
    ppt_path = test_data_dir / "test.pptx"
    out_pdf = tmp_path / "ppt.pdf"
    res_pdf = convert_to_pdf(ppt_path, output_path=out_pdf)
    assert res_pdf.exists()
    assert res_pdf.stat().st_size > 0

    with pymupdf.open(str(res_pdf)) as pdf:
        assert len(pdf) == 2
        text_slide1 = pdf[0].get_text()
        text_slide2 = pdf[1].get_text()
        assert "Operational Efficiency" in text_slide1
        assert "Core Objectives" in text_slide2


def test_convert_word_fallback(test_data_dir: Path, tmp_path: Path):
    doc_path = test_data_dir / "test.docx"
    out_pdf = tmp_path / "word_fallback.pdf"
    success = convert_word_fallback(doc_path, out_pdf)
    assert success
    assert out_pdf.exists()
    assert out_pdf.stat().st_size > 0
    with pymupdf.open(str(out_pdf)) as pdf:
        assert len(pdf) >= 1


def test_convert_excel_merged_cells(tmp_path: Path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MergedReport"
    ws.merge_cells("A1:C1")
    ws["A1"] = "Consolidated Report 2026"
    ws.append(["Metric A", "Metric B", "Metric C"])
    ws.append(["100", "200", "300"])
    excel_path = tmp_path / "merged.xlsx"
    wb.save(str(excel_path))

    out_pdf = tmp_path / "merged.pdf"
    success = convert_excel_fallback(excel_path, out_pdf)
    assert success
    assert out_pdf.exists()
    with pymupdf.open(str(out_pdf)) as pdf:
        assert "Consolidated Report 2026" in pdf[0].get_text()
