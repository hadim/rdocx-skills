"""docx, creating a document from nothing and reading it back with python-docx (an independent reader)."""
import io

import rdocx
from PIL import Image



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


def test_add_picture_with_one_dimension_keeps_the_aspect_ratio():
    """As python-docx: a width alone (or a height alone) scales the other side."""
    d = rdocx.Document()
    png = io.BytesIO()
    Image.new("RGB", (20, 10)).save(png, "PNG")
    d.add_picture(png.getvalue(), "x.png", width=rdocx.Inches(2))
    d.add_picture(png.getvalue(), height=rdocx.Inches(2))
    assert [(p.width, p.height) for p in d.pictures] == [(rdocx.Inches(2), rdocx.Inches(1)), (rdocx.Inches(4), rdocx.Inches(2))]
