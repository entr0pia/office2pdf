"""
Unit tests for pdf_to_images.py.
"""

from pathlib import Path
import pytest
from PIL import Image

from scripts.convert_office import convert_to_pdf
from scripts.pdf_to_images import pdf_to_images, parse_page_selection
from tests.generate_samples import generate_sample_powerpoint


@pytest.fixture(scope="session")
def sample_pdf(tmp_path_factory) -> Path:
    d = tmp_path_factory.mktemp("pdf_data")
    pptx_path = d / "test.pptx"
    generate_sample_powerpoint(pptx_path)
    pdf_path = d / "test.pdf"
    convert_to_pdf(pptx_path, output_path=pdf_path)
    return pdf_path


def test_parse_page_selection():
    assert parse_page_selection(None, 5) == {1, 2, 3, 4, 5}
    assert parse_page_selection("all", 5) == {1, 2, 3, 4, 5}
    assert parse_page_selection("1", 5) == {1}
    assert parse_page_selection("1,3", 5) == {1, 3}
    assert parse_page_selection("2-4", 5) == {2, 3, 4}
    assert parse_page_selection("1,3-5", 5) == {1, 3, 4, 5}


def test_pdf_to_images_all_pages(sample_pdf: Path, tmp_path: Path):
    out_dir = tmp_path / "images_all"
    images = pdf_to_images(sample_pdf, output_dir=out_dir, dpi=150, fmt="png")
    assert len(images) == 2
    for img_path in images:
        assert img_path.exists()
        assert img_path.stat().st_size > 0
        with Image.open(str(img_path)) as img:
            assert img.format == "PNG"
            assert img.width > 500
            assert img.height > 300


def test_pdf_to_images_single_page(sample_pdf: Path, tmp_path: Path):
    out_dir = tmp_path / "images_single"
    images = pdf_to_images(sample_pdf, output_dir=out_dir, pages="2", dpi=200, fmt="jpg")
    assert len(images) == 1
    assert "002" in images[0].name
    assert images[0].suffix == ".jpg"
    with Image.open(str(images[0])) as img:
        assert img.format == "JPEG"
