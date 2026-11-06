"""
Unit tests for inspect_office.py.
"""

import json
from pathlib import Path
import pytest

from scripts.inspect_office import (
    inspect_file,
    inspect_word,
    inspect_excel,
    inspect_powerpoint,
    format_inspection_output,
)
from tests.generate_samples import (
    generate_sample_word,
    generate_sample_excel,
    generate_sample_powerpoint,
)


@pytest.fixture(scope="session")
def test_data_dir(tmp_path_factory) -> Path:
    d = tmp_path_factory.mktemp("office_data")
    generate_sample_word(d / "test.docx")
    generate_sample_excel(d / "test.xlsx")
    generate_sample_powerpoint(d / "test.pptx")
    return d


def test_inspect_word(test_data_dir: Path):
    doc_path = test_data_dir / "test.docx"
    res = inspect_word(doc_path)
    assert res["file_type"] == "word"
    assert res["paragraph_count"] >= 4
    assert res["heading_count"] == 3
    assert res["headings"][0]["text"] == "Project Overview Report"
    assert res["headings"][1]["text"] == "1. Executive Summary"
    assert res["table_count"] == 1
    assert res["tables"][0]["rows"] == 3
    assert res["tables"][0]["columns"] == 3
    assert "Scenario" in res["tables"][0]["headers"]


def test_inspect_excel(test_data_dir: Path):
    excel_path = test_data_dir / "test.xlsx"
    res = inspect_excel(excel_path)
    assert res["file_type"] == "excel"
    assert res["sheet_count"] == 2
    sheet_names = [s["name"] for s in res["sheets"]]
    assert "Metrics" in sheet_names
    assert "Reference Data" in sheet_names

    metrics_sheet = res["sheets"][0]
    assert metrics_sheet["max_rows"] >= 4
    assert len(metrics_sheet["preview_rows"]) >= 4
    assert metrics_sheet["preview_rows"][0][0] == "ID"


def test_inspect_powerpoint(test_data_dir: Path):
    ppt_path = test_data_dir / "test.pptx"
    res = inspect_powerpoint(ppt_path)
    assert res["file_type"] == "powerpoint"
    assert res["slide_count"] == 2
    assert "Operational Efficiency Strategy" in res["slides"][0]["title"]
    assert "Core Objectives" in res["slides"][1]["title"]
    assert "multimodal LLM" in res["slides"][1]["notes"]


def test_format_inspection_output(test_data_dir: Path):
    excel_path = test_data_dir / "test.xlsx"
    res = inspect_excel(excel_path)
    formatted = format_inspection_output(res)
    assert "=== Document Inspection: test.xlsx (EXCEL) ===" in formatted
    assert "Sheet: [Metrics]" in formatted


def test_inspect_file_dispatcher(test_data_dir: Path):
    res_docx = inspect_file(test_data_dir / "test.docx")
    assert res_docx["file_type"] == "word"

    res_xlsx = inspect_file(test_data_dir / "test.xlsx")
    assert res_xlsx["file_type"] == "excel"

    res_pptx = inspect_file(test_data_dir / "test.pptx")
    assert res_pptx["file_type"] == "powerpoint"


def test_inspect_word_chinese_headings(tmp_path: Path):
    import docx
    doc = docx.Document()
    doc.add_heading("两率一效工作方案", level=0)
    p1 = doc.add_paragraph("正文内容")
    p1.style = doc.styles["Normal"]
    h1 = doc.add_paragraph("一、总体目标")
    h1.style = doc.styles["Heading 1"]
    h2 = doc.add_paragraph("1.1 业务指标")
    h2.style = doc.styles["Heading 2"]

    doc_path = tmp_path / "chinese_doc.docx"
    doc.save(str(doc_path))

    res = inspect_word(doc_path)
    assert res["heading_count"] >= 2
    headings_text = [h["text"] for h in res["headings"]]
    assert "两率一效工作方案" in headings_text or "一、总体目标" in headings_text


def test_inspect_legacy_formats(tmp_path: Path):
    from scripts.inspect_office import inspect_legacy_office, format_inspection_output
    dummy_doc = tmp_path / "legacy_test.doc"
    dummy_doc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")  # OLE header magic bytes
    res = inspect_file(dummy_doc)
    assert res["file_type"] == "legacy_office"
    assert res["extension"] == ".doc"

    formatted = format_inspection_output(res)
    assert "Legacy Binary Office (.doc)" in formatted
    assert "Notice:" in formatted


def test_inspect_nonexistent_file():
    with pytest.raises(FileNotFoundError):
        inspect_file("nonexistent_file_12345.docx")


def test_inspect_unsupported_format(tmp_path: Path):
    dummy = tmp_path / "test.unknown"
    dummy.write_text("dummy")
    with pytest.raises(ValueError, match="Unsupported file format"):
        inspect_file(dummy)
