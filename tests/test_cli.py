"""
Unit tests for CLI interface (scripts/office2pdf.py).
"""

import os
import subprocess
import sys
from pathlib import Path
import pytest

from tests.generate_samples import generate_sample_word, generate_sample_excel


@pytest.fixture(scope="session")
def cli_env() -> dict:
    env = os.environ.copy()
    skill_path = str(Path(__file__).resolve().parent.parent / "skills" / "office2pdf")
    current_pp = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = f"{skill_path}{os.pathsep}{current_pp}" if current_pp else skill_path
    return env


@pytest.fixture(scope="session")
def cli_data_dir(tmp_path_factory) -> Path:
    d = tmp_path_factory.mktemp("cli_data")
    generate_sample_word(d / "cli_sample.docx")
    generate_sample_excel(d / "cli_sample.xlsx")
    return d


def test_cli_inspect_text(cli_data_dir: Path, cli_env: dict):
    doc_path = cli_data_dir / "cli_sample.docx"
    cmd = [sys.executable, "-m", "scripts.office2pdf", "inspect", str(doc_path)]
    res = subprocess.run(cmd, capture_output=True, text=True, check=True, env=cli_env)
    assert "Document Inspection: cli_sample.docx (WORD)" in res.stdout
    assert "Project Overview Report" in res.stdout


def test_cli_inspect_json(cli_data_dir: Path, cli_env: dict):
    excel_path = cli_data_dir / "cli_sample.xlsx"
    cmd = [sys.executable, "-m", "scripts.office2pdf", "inspect", str(excel_path), "--json"]
    res = subprocess.run(cmd, capture_output=True, text=True, check=True, env=cli_env)
    assert '"file_type": "excel"' in res.stdout
    assert '"Metrics"' in res.stdout


def test_cli_convert_pdf(cli_data_dir: Path, tmp_path: Path, cli_env: dict):
    doc_path = cli_data_dir / "cli_sample.docx"
    out_pdf = tmp_path / "cli_word.pdf"
    cmd = [sys.executable, "-m", "scripts.office2pdf", "pdf", str(doc_path), "-o", str(out_pdf)]
    res = subprocess.run(cmd, capture_output=True, text=True, check=True, env=cli_env)
    assert "SUCCESS:" in res.stdout
    assert out_pdf.exists()
    assert out_pdf.stat().st_size > 0


def test_cli_image(cli_data_dir: Path, tmp_path: Path, cli_env: dict):
    excel_path = cli_data_dir / "cli_sample.xlsx"
    out_dir = tmp_path / "cli_images"
    cmd = [sys.executable, "-m", "scripts.office2pdf", "image", str(excel_path), "-o", str(out_dir), "--dpi", "150"]
    res = subprocess.run(cmd, capture_output=True, text=True, check=True, env=cli_env)
    assert "SUCCESS: Rendered" in res.stdout
    assert out_dir.exists()
    imgs = list(out_dir.glob("*.png"))
    assert len(imgs) >= 1
    # Verify no clutter left in source folder
    assert not (cli_data_dir / "cli_sample.pdf").exists()
