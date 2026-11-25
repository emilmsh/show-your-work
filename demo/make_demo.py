"""Builds the Office files and the PDF of the demo set.

The text files beside this script (notes.md, analysis.R) are written by hand.
No Office needed:

    pip install python-docx openpyxl python-pptx pymupdf
    python demo/make_demo.py

An earlier version drove Word, Excel and PowerPoint themselves and hung in
Word, both in a running Word and in one of its own. Writing the files
directly cannot hang, and runs anywhere.
"""

from pathlib import Path

import fitz
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font
from pptx import Presentation

HERE = Path(__file__).parent


def report() -> None:
    doc = Document()
    doc.styles["Normal"].font.size = Pt(12)
    doc.add_heading("School meals pilot: evaluation 2025", level=0)

    doc.add_heading("1 Summary", level=1)
    doc.add_paragraph(
        "In 2023 forty primary schools began serving a free lunch every school "
        "day. This report evaluates what the pilot did for attendance and what "
        "it cost."
    )
    doc.add_paragraph(
        "Attendance rose by 4.2 percentage points in the pilot schools, compared "
        "with matched schools that did not take part. The effect is largest for "
        "pupils in the lowest income quartile."
    )
    doc.add_paragraph("The pilot reached 1,987 pupils and cost 4,800 kroner per pupil per year.")
    doc.add_page_break()

    doc.add_heading("2 Data", level=1)
    doc.add_paragraph(
        "We use the national attendance register for the school years 2019 to "
        "2025. The sample is 40 pilot schools and 120 matched control schools."
    )
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    for row, cells in zip(
        table.rows,
        (("Group", "Schools", "Pupils"), ("Pilot", "40", "1,987"), ("Control", "120", "6,214")),
    ):
        for cell, text in zip(row.cells, cells):
            cell.text = text
    doc.add_paragraph("")
    doc.add_paragraph(
        "Schools were matched on size, region and attendance in the three years before the pilot."
    )
    doc.add_page_break()

    doc.add_heading("3 Results", level=1)
    doc.add_paragraph(
        "The main finding holds for alternative specifications, and it remains "
        "when the first pilot year is left out."
    )
    doc.add_paragraph(
        "We find no effect on test scores in the first two years. Two years is a "
        "short time for an effect on learning to show."
    )
    doc.add_paragraph("The total cost of the pilot was 9.5 million kroner per year.")
    doc.save(HERE / "report.docx")


def slides() -> None:
    deck = Presentation()
    layout = deck.slide_layouts[1]  # a title and a body
    pages = (
        ("School meals pilot", ("Forty schools, a free lunch every school day.", "Evaluation of the first two years.")),
        ("Method", ("Attendance register, 2019 to 2025.", "Difference-in-differences with matched control schools.")),
        ("Results", ("Attendance rose by 4.2 percentage points.", "No effect on test scores so far.")),
        ("Recommendation", ("Continue the pilot for two more years.", "Measure learning again in 2027.")),
    )
    for title, lines in pages:
        slide = deck.slides.add_slide(layout)
        slide.shapes.title.text = title
        body = slide.placeholders[1].text_frame
        body.text = lines[0]
        for line in lines[1:]:
            body.add_paragraph().text = line
    deck.slides[3].notes_slide.notes_text_frame.text = (
        "Mention that the cost is 4,800 kroner per pupil per year."
    )
    deck.save(HERE / "slides.pptx")


def model() -> None:
    book = Workbook()
    assumptions = book.active
    assumptions.title = "Assumptions"
    for row in (
        ("Parameter", "Value"),
        ("Pupils in the pilot", 1987),
        ("Cost per pupil per year (kroner)", 4800),
        ("School days per year", 190),
    ):
        assumptions.append(row)
    assumptions.column_dimensions["A"].width = 36

    results = book.create_sheet("Results")
    for row in (
        ("Measure", "Value"),
        ("Total cost per year (kroner)", "=Assumptions!B2*Assumptions!B3"),
        ("Cost per pupil per school day (kroner)", "=Assumptions!B3/Assumptions!B4"),
        ("Effect on attendance (percentage points)", 4.2),
        ("Cost per percentage point (kroner)", "=B2/B4"),
    ):
        results.append(row)
    results.column_dimensions["A"].width = 42
    results.column_dimensions["B"].width = 16
    for sheet in (assumptions, results):
        for cell in sheet[1]:
            cell.font = Font(bold=True)
    book.active = 1
    book.save(HERE / "model.xlsx")


def appendix() -> None:
    pdf = fitz.open()
    pages = (
        (
            "Technical appendix",
            "A.1 Estimation",
            (
                "The effect is estimated by difference-in-differences, comparing the change in "
                "attendance in pilot schools with the change in matched control schools.",
                "School and year fixed effects are included in every specification.",
            ),
        ),
        (
            "",
            "A.2 Inference",
            (
                "Standard errors are clustered at the school level, since attendance is correlated "
                "among pupils of the same school.",
                "The pre-pilot trends of the two groups are parallel: none of the six pre-period "
                "coefficients differs from zero at the five per cent level.",
            ),
        ),
    )
    for title, heading, paragraphs in pages:
        page = pdf.new_page(width=595, height=842)
        y = 90.0
        if title:
            page.insert_text((72, y), title, fontname="tibo", fontsize=22)
            y += 44
        page.insert_text((72, y), heading, fontname="tibo", fontsize=15)
        y += 18
        for text in paragraphs:
            box = fitz.Rect(72, y, 523, y + 90)
            page.insert_textbox(box, text, fontname="tiro", fontsize=12, lineheight=1.4)
            y += 72
    pdf.save(HERE / "appendix.pdf")


if __name__ == "__main__":
    report()
    slides()
    model()
    appendix()
    for name in ("report.docx", "slides.pptx", "model.xlsx", "appendix.pdf"):
        print(name, (HERE / name).stat().st_size)
