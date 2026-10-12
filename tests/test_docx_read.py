"""docx, reading: text, structure, JSON views, metadata, hyperlinks, sections, styles."""
import json
import re
import zipfile

import docx
import rdocx
from docx.oxml.ns import qn
from docx.shared import Twips

from conftest import run


# ---------------------------------------------------------------- validation
def broken(src, dst, name, change):
    with zipfile.ZipFile(src) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(dst, "w") as z:
        for info, data in items:
            z.writestr(info, change(data) if info.filename == name else data)
    return dst


# ---------------------------------------------------------------- measurements with a decimal part
def measured(path):
    """An indented paragraph with spacing, a 1x2 table with a row height and a width, a header with an indent: the
    measurements Google Docs writes with a floating-point tail."""
    d = docx.Document()
    p = d.add_paragraph("Indented paragraph")
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Twips(720), Twips(-227)
    p.paragraph_format.space_before, p.paragraph_format.line_spacing = Twips(120), 1.15
    t = d.add_table(rows=1, cols=2)
    t.rows[0].height = Twips(535)
    t.cell(0, 0).text = "Cell A"
    width = t._tbl.tblPr.find(qn("w:tblW"))
    width.set(qn("w:type"), "dxa")
    width.set(qn("w:w"), "8640")
    d.sections[0].header.paragraphs[0].text = "Header"
    d.sections[0].header.paragraphs[0].paragraph_format.left_indent = Twips(360)
    d.save(path)
    return path


def twips(emu):
    return emu / 635


# part, attribute as python-docx writes it with {} for the value, a decimal value, the integer it stands for, and
# where the binding reads it back in twips (None: not exposed)
DECIMAL = [
    ("document", 'w:gridCol w:w="{}"', "4320.0", 4320, lambda d: twips(d.tables[0].grid_widths[0])),
    ("document", 'w:gridCol w:w="{}"', "4319.999999999999", 4320, lambda d: twips(d.tables[0].grid_widths[0])),
    ("document", 'w:left="{}"', "720.0", 720, lambda d: twips(d.paragraphs[0].paragraph_format.left_indent)),
    ("document", 'w:hanging="{}"', "226.99999999999977", 227, lambda d: -twips(d.paragraphs[0].paragraph_format.first_line_indent)),
    ("document", 'w:before="{}"', "120.0", 120, lambda d: twips(d.paragraphs[0].paragraph_format.space_before)),
    ("document", 'w:trHeight w:val="{}"', "535.0000000000182", 535, lambda d: twips(d.tables[0].rows[0].height)),
    ("document", 'w:pgSz w:w="{}"', "12240.0", 12240, lambda d: twips(d.sections[0].page_width)),
    ("document", 'w:top="{}"', "1440.0", 1440, lambda d: twips(d.sections[0].margin_top)),
    ("document", 'w:tcW w:type="dxa" w:w="{}"', "4319.999999999999", 4320, lambda d: twips(d.tables[0].cell(0, 0).width)),
    ("document", 'w:tblW w:type="dxa" w:w="{}"', "8639.999999999999", 8640, lambda d: twips(d.tables[0].width)),
    ("styles", 'w:after="{}"', "200.0", 200, None),
    ("styles", '<w:sz w:val="{}"/>', "22.0", 22, None),
    ("numbering", 'w:left="{}"', "360.0", 360, None),
]


def test_a_tab_reads_as_a_tab_in_every_story_view(rdocx_cli, tmp_path):
    """A w:tab between two runs, as Word writes a footer "title<tab>page": every text view keeps it."""
    d = docx.Document()
    for p in (d.add_paragraph(), d.sections[0].header.paragraphs[0], d.sections[0].footer.paragraphs[0]):
        p.add_run("left")
        p.add_run().add_tab()
        p.add_run("right")
    d.save(tmp_path / "t.docx")
    doc = rdocx.Document.open(tmp_path / "t.docx")
    assert doc.sections[0].footer.paragraphs[0].text == "left\tright"  # the handle keeps it already
    assert {it.text for it in doc.story_items if it.kind == "paragraph"} == {"left\tright"}
    plain = run([rdocx_cli, "text", tmp_path / "t.docx"], check=True).stdout
    assert plain.count("left\tright") == 3
    record = json.loads(run([rdocx_cli, "text", "--json", tmp_path / "t.docx"], check=True).stdout)
    assert {i["text"] for s in record["stories"] for i in s["items"]} == {"left\tright"}


def test_a_break_reads_as_in_paragraph_text_and_set_story_text_writes_tabs_and_breaks(tmp_path):
    """A line break in a header run reads as a newline in StoryItem.text, as in Paragraph.text; set_story_text
    writes a tab and a newline as a w:tab and a w:br, never as literal characters in w:t."""
    d = docx.Document()
    p = d.sections[0].header.paragraphs[0]
    p.add_run("line one")
    p.add_run().add_break()
    p.add_run("line two")
    d.save(tmp_path / "b.docx")
    doc = rdocx.Document.open(tmp_path / "b.docx")
    item = next(it for it in doc.story_items if it.story.kind == "header" and it.kind == "paragraph")
    assert item.text == doc.sections[0].header.paragraphs[0].text == "line one\nline two"
    doc.set_story_text(item, "Title\tPage\nSecond line")
    doc.save(tmp_path / "s.docx")
    doc = rdocx.Document.open(tmp_path / "s.docx")
    assert doc.sections[0].header.paragraphs[0].text == "Title\tPage\nSecond line"
    header = next(n for n in zipfile.ZipFile(tmp_path / "s.docx").namelist() if re.match(r"word/header\d+\.xml$", n))
    xml = zipfile.ZipFile(tmp_path / "s.docx").read(header).decode()
    assert "<w:tab/>" in xml and "<w:br/>" in xml and "\t" not in xml
