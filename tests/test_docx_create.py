"""docx, creating a document from nothing and reading it back with python-docx (an independent reader)."""
import io

import docx
import pytest
import rdocx
from PIL import Image

from conftest import run


def build(path):
    d = rdocx.Document()
    d.add_paragraph("Quarterly note")
    d.paragraphs[0].style = "Heading1"
    d.add_paragraph("Body text ")
    d.paragraphs[1].add_run("in bold")
    f = d.paragraphs[1].runs[1].font
    f.bold = True
    d.paragraphs[1].runs[1].font.color = rdocx.RGBColor(0x7B, 0x1E, 0x3A)
    d.paragraphs[1].runs[1].font.size = rdocx.Pt(14)
    d.paragraphs[1].runs[1].font.name = "Arial"
    d.paragraphs[1].runs[1].font.highlight = "yellow"
    d.paragraphs[1].runs[1].font.underline = rdocx.WD_UNDERLINE.SINGLE
    pf = d.paragraphs[1].paragraph_format
    pf.space_after = rdocx.Pt(12)
    d.paragraphs[1].paragraph_format.line_spacing = 1.5
    d.paragraphs[1].paragraph_format.keep_with_next = True
    d.paragraphs[1].alignment = rdocx.WD_ALIGN_PARAGRAPH.JUSTIFY
    d.add_table(2, 3)
    d.tables[0].cell(0, 0).text = "A"
    d.tables[0].cell(1, 2).text = "Z"
    d.tables[0].alignment = rdocx.WD_TABLE_ALIGNMENT.CENTER
    d.set_header("Header text")
    d.set_footer("Footer text")
    png = io.BytesIO()
    Image.new("RGB", (200, 100), (10, 200, 10)).save(png, "PNG")
    d.add_picture(png.getvalue(), "green.png", width=rdocx.Inches(2), height=rdocx.Inches(1))
    d.paragraphs[0].add_hyperlink(" (source)", "https://example.org/source")
    d.save(path)
    return path


def test_create_and_read_back_with_python_docx(tmp_path, rdocx_cli):
    path = build(tmp_path / "new.docx")
    x = docx.Document(path)
    assert x.paragraphs[0].style.name == "Heading 1"
    assert x.paragraphs[0].text == "Quarterly note (source)"
    run_ = x.paragraphs[1].runs[1]
    assert run_.bold and str(run_.font.color.rgb) == "7B1E3A" and run_.font.size.pt == 14
    assert run_.font.name == "Arial" and str(run_.font.highlight_color).startswith("YELLOW")
    assert x.paragraphs[1].paragraph_format.line_spacing == 1.5
    assert x.tables[0].cell(1, 2).text == "Z"
    assert x.sections[0].header.paragraphs[0].text == "Header text"
    assert len(x.inline_shapes) == 1 and round(x.inline_shapes[0].width.inches, 2) == 2.0
    res = run([rdocx_cli, "validate", path])
    assert res.returncode == 0, res.stdout + res.stderr


def test_new_document_has_normal_and_heading1():
    assert {"Normal", "Heading1"} <= {s.style_id for s in rdocx.Document().styles}


def test_new_document_has_the_usual_styles():
    """Today a new Document holds only Normal and Heading1: other styles come from a template file."""
    assert {"Heading2", "Heading9", "Title", "Subtitle", "NoSpacing", "Quote", "ListParagraph", "Caption",
            "TableGrid"} <= {s.style_id for s in rdocx.Document().styles}


def test_new_document_has_no_title_or_author(rdocx_cli, tmp_path):
    """A new document has no title or author (set them with `core_properties`): `validate` warns, exit 0."""
    rdocx.Document().save(tmp_path / "n.docx")
    res = run([rdocx_cli, "validate", tmp_path / "n.docx"])
    assert res.returncode == 0 and "title" in (res.stdout + res.stderr).lower()


def test_add_picture_with_one_dimension_keeps_the_aspect_ratio():
    """As python-docx: a width alone (or a height alone) scales the other side."""
    d = rdocx.Document()
    png = io.BytesIO()
    Image.new("RGB", (20, 10)).save(png, "PNG")
    d.add_picture(png.getvalue(), "x.png", width=rdocx.Inches(2))
    d.add_picture(png.getvalue(), height=rdocx.Inches(2))
    assert [(p.width, p.height) for p in d.pictures] == [(rdocx.Inches(2), rdocx.Inches(1)), (rdocx.Inches(4), rdocx.Inches(2))]
