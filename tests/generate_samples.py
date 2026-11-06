"""
Generate sample Office test documents: Word (.docx), Excel (.xlsx), PowerPoint (.pptx).
"""

from pathlib import Path


def generate_sample_word(path: Path) -> Path:
    import docx

    doc = docx.Document()
    doc.add_heading("Project Overview Report", level=0)
    doc.add_paragraph("This is an automated test document for office2pdf verification.")

    doc.add_heading("1. Executive Summary", level=1)
    doc.add_paragraph("Multimodal LLMs require crisp visual representations of complex documents.")

    doc.add_heading("2. Evaluation Metrics", level=2)
    doc.add_paragraph("Key metrics include penetration rate, accuracy, and operational efficiency.")

    # Add a table
    table = doc.add_table(rows=3, cols=3)
    headers = ["Scenario", "Metric Name", "Target Value"]
    for c_idx, h in enumerate(headers):
        table.cell(0, c_idx).text = h

    data = [
        ["Substation Inspection", "Coverage Rate", "98.5%"],
        ["Defect Recognition", "Accuracy Rate", "95.2%"],
    ]
    for r_idx, row in enumerate(data, start=1):
        for c_idx, val in enumerate(row):
            table.cell(r_idx, c_idx).text = val

    doc.save(str(path))
    return path


def generate_sample_excel(path: Path) -> Path:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment

    wb = openpyxl.Workbook()

    # Sheet 1: Metrics
    ws1 = wb.active
    ws1.title = "Metrics"
    ws1.append(["ID", "Scenario Name", "Target Penetration", "Accuracy", "Efficiency Boost"])
    ws1.append([1, "Substation Inspection", "90%", "95%", "+30%"])
    ws1.append([2, "Transmission Drone Inspection", "85%", "92%", "+45%"])
    ws1.append([3, "Smart Customer Service", "99%", "98%", "+60%"])

    # Style header
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    for col in range(1, 6):
        cell = ws1.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    ws1.cell(row=2, column=3).fill = yellow_fill

    # Sheet 2: Reference
    ws2 = wb.create_sheet(title="Reference Data")
    ws2.append(["Code", "Department", "Contact Person"])
    ws2.append(["MKT-01", "Marketing Dept", "Alice"])
    ws2.append(["DEV-02", "Equipment Dept", "Bob"])

    wb.save(str(path))
    return path


def generate_sample_powerpoint(path: Path) -> Path:
    import pptx
    from pptx.util import Inches, Pt

    prs = pptx.Presentation()

    # Slide 1: Title slide
    slide_layout_title = prs.slide_layouts[0]
    slide1 = prs.slides.add_slide(slide_layout_title)
    slide1.shapes.title.text = "Operational Efficiency Strategy"
    slide1.placeholders[1].text = "AI-Driven Document Processing & Vision Verification"

    # Slide 2: Bullet slide
    slide_layout_bullets = prs.slide_layouts[1]
    slide2 = prs.slides.add_slide(slide_layout_bullets)
    slide2.shapes.title.text = "Core Objectives"
    tf = slide2.placeholders[1].text_frame
    tf.text = "1. High fidelity visual perception via PDF"
    p2 = tf.add_paragraph()
    p2.text = "2. Zero-token code-based metadata inspection"
    p3 = tf.add_paragraph()
    p3.text = "3. Seamless multi-engine fallback"

    # Add speaker notes
    notes = slide2.notes_slide.notes_text_frame
    notes.text = "Discuss multimodal LLM advantages for chart reading."

    prs.save(str(path))
    return path


if __name__ == "__main__":
    test_dir = Path("test_data")
    test_dir.mkdir(exist_ok=True)
    generate_sample_word(test_dir / "sample.docx")
    generate_sample_excel(test_dir / "sample.xlsx")
    generate_sample_powerpoint(test_dir / "sample.pptx")
    print("Test samples generated in test_data/")
