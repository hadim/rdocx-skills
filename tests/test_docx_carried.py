"""docx: fixes the pinned integration build (upstream v0.16.4 plus tensorbee/rdocx#427, #428 and #438 to #444, #443
also carrying the cell setters of #445) carries and upstream `main` does not have yet, which the skills rely on, and
the cases of those issues that stay open as known gaps (strict xfails, tests/gaps.py). Each test names the pull
request that carries it, or the issue of its gap: delete the test once upstream main has it. Each test builds its own
input."""
import functools
import io
import json
import re
import struct
import time
import warnings
import xml.etree.ElementTree as ET
import zipfile
import zlib
from pathlib import Path

import pytest
import rdocx
from rdocx.enum.table import WD_ALIGN_VERTICAL
from rdocx.enum.text import WD_ALIGN_PARAGRAPH
from rdocx.shared import Pt

from conftest import pdf_uses_font, run


def rewritten(doc, edit):
    """doc saved, its parts passed to edit(parts) (a dict of name: bytes), and opened again."""
    archive = zipfile.ZipFile(io.BytesIO(doc.to_bytes()))
    parts = {name: archive.read(name) for name in archive.namelist()}
    edit(parts)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return rdocx.Document.from_bytes(out.getvalue())


def body_top(header_lines):
    """The y of the first body line on page 1, under a 2 cm top margin and a header of header_lines lines from
    1.27 cm."""
    doc = rdocx.Document()
    doc.update_section(0, margin_top=rdocx.Cm(2.0), header_distance=rdocx.Cm(1.27))
    for i in range(header_lines):
        doc.section(0).header.add_paragraph(f"Header line {i + 1}")
    for i in range(40):
        doc.add_paragraph(f"Body line {i + 1}.")
    fragments = doc.layout()
    first = min(f.physical_page for f in fragments)
    return min(f.bounds.y for f in fragments if f.physical_page == first)


# carried by tensorbee/rdocx#428 (#424), delete once upstream main has it
def test_body_starts_below_a_header_taller_than_the_top_margin():
    """A header that runs below the top margin moves the body down, as in Word, so renders, PDFs and the page
    numbers of a rebuilt TOC match Word's. The body used to start at the top margin, under the header's lines
    (tensorbee/rdocx#424)."""
    one, five = body_top(1), body_top(5)
    assert one == pytest.approx(2.0 / 2.54 * 72, abs=0.5)
    assert five > one + 3 * 10


def toc_document(merged, instruction):
    """Two Heading 1 paragraphs, a table between them (its first column merged vertically over two rows when
    merged) and a TOC field whose instruction is replaced by instruction."""
    doc = rdocx.Document()
    doc.add_paragraph("Title page")
    doc.add_paragraph("First section")
    doc.paragraphs[-1].style = "Heading 1"
    doc.add_table(3, 2, style="Table Grid")
    for r in range(3):
        for c in range(2):
            doc.tables[-1].cell(r, c).text = f"cell {r}{c}"
    if merged:
        doc.tables[-1].set_cell_vertical_merge(0, 0, "restart")
        doc.tables[-1].set_cell_vertical_merge(1, 0, "continue")
    doc.add_paragraph("Second section")
    doc.paragraphs[-1].style = "Heading 1"
    doc.insert_toc(1, max_level=2)

    def edit(parts):
        xml, n = re.subn(r"(<w:instrText[^>]*>)\s*TOC[^<]*(</w:instrText>)",
                         lambda m: m.group(1) + instruction + m.group(2), parts["word/document.xml"].decode())
        assert n == 1
        parts["word/document.xml"] = xml.encode()

    return rewritten(doc, edit)


# carried by tensorbee/rdocx#428 (#426), delete once upstream main has it
@pytest.mark.parametrize("merged", [False, True])
def test_rebuild_toc_with_u_switch_beside_a_vertical_merge(merged):
    """A TOC whose instruction carries \\u, as Word writes it for many templates, rebuilds in a document whose table
    has a vertically merged cell (python-api.md, rebuild_toc). It used to raise RdocxError ("TOC outline source has no
    qualified current layout facts", tensorbee/rdocx#426)."""
    doc = toc_document(merged, ' TOC \\o "1-2" \\h \\u ')
    assert doc.rebuild_toc().entry_count == 2


# carried by tensorbee/rdocx#427 (#425), delete once upstream main has it
def test_compare_keeps_the_revisions_of_a_redline():
    """A second reviewer's edit of a redline compares against it: the first reviewer's revisions stay as they are and
    the edit adds revisions by the second (python-api.md, compare). It used to raise RdocxError ("document comparison
    requires inputs without existing modeled revisions", tensorbee/rdocx#425)."""
    v1 = rdocx.Document()
    v1.add_paragraph("The first sentence stays. The second sentence changes once.")
    v1.add_paragraph("Another paragraph changes later.")
    v1 = rdocx.Document.from_bytes(v1.to_bytes())
    v2 = rdocx.Document.from_bytes(v1.to_bytes())
    v2.paragraphs[0].replace_text("changes once", "changed once", expect=1)
    redline = rdocx.Document.from_bytes(v1.to_bytes())
    redline.compare(v2, "Reviewer A", "2026-10-27T12:00:00Z", granularity="word")
    redline = rdocx.Document.from_bytes(redline.to_bytes())
    first = sorted((r.author, r.kind) for r in redline.revisions)
    edited = rdocx.Document.from_bytes(redline.to_bytes())
    edited.paragraphs[1].replace_text("changes later", "changed later", expect=1)
    second = rdocx.Document.from_bytes(redline.to_bytes())
    second.compare(edited, "Reviewer B", "2026-10-27T13:00:00Z", granularity="word")
    authors = [r.author for r in second.revisions]
    assert authors.count("Reviewer A") == len(first) and authors.count("Reviewer B") >= 1


# ---- tensorbee/rdocx#441 (#432) and #440 (#433) ----------------------------------------------------------------------


def flat(xml):
    """xml without the whitespace between tags."""
    return re.sub(r">\s+<", "><", xml)


def patched(doc, edits):
    """doc saved, each part named in edits (name: function of its text returning the new text) rewritten, and opened
    again."""
    archive = zipfile.ZipFile(io.BytesIO(doc.to_bytes()))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in archive.infolist():
            data = archive.read(item)
            if item.filename in edits:
                data = edits[item.filename](data.decode()).encode()
            z.writestr(item, data)
    return rdocx.Document.from_bytes(out.getvalue())


def png_1x1():
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\xff\xff\xff")) + chunk(b"IEND", b""))


def decode_png(data):
    """Minimal decoder for 8-bit RGB or RGBA non-interlaced PNG: (width, height, rows of (r, g, b))."""
    pos, idat = 8, b""
    while pos < len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        chunk = data[pos + 8:pos + 8 + length]
        if kind == b"IHDR":
            width, height, depth, ctype = struct.unpack(">IIBB", chunk[:10])
            assert depth == 8 and ctype in (2, 6)
            bpp = 3 if ctype == 2 else 4
        elif kind == b"IDAT":
            idat += chunk
        pos += 12 + length
    raw, stride, rows, prev = zlib.decompress(idat), width * bpp, [], bytearray(width * bpp)
    for y in range(height):
        f, line = raw[y * (stride + 1)], bytearray(raw[y * (stride + 1) + 1:(y + 1) * (stride + 1)])
        for i in range(stride):
            a = line[i - bpp] if i >= bpp else 0
            b, c = prev[i], prev[i - bpp] if i >= bpp else 0
            if f == 1:
                line[i] = (line[i] + a) & 255
            elif f == 2:
                line[i] = (line[i] + b) & 255
            elif f == 3:
                line[i] = (line[i] + (a + b) // 2) & 255
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append([tuple(line[x * bpp:x * bpp + 3]) for x in range(width)])
        prev = line
    return width, height, rows



EQUATIONS = {
    "latex sqrt": (r"\sqrt{x}", "latex"),
    "latex cube root": (r"\sqrt[3]{x}", "latex"),
    "mathml msqrt": ('<math xmlns="http://www.w3.org/1998/Math/MathML"><msqrt><mi>x</mi></msqrt></math>', "mathml"),
}


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("source", EQUATIONS)
def test_add_equation_writes_the_degree_of_a_root(source):
    """Every m:rad that add_equation writes holds an m:deg after m:radPr, empty for a square root, as CT_Rad requires
    (python-api.md, add_equation). A square root used to have no m:deg (tensorbee/rdocx#432)."""
    src, fmt = EQUATIONS[source]
    doc = rdocx.Document()
    doc.add_paragraph("").add_equation(src, format=fmt)
    rads = re.findall(r"<m:rad>.*?</m:rad>", doc.part_xml("word/document.xml"), re.S)
    assert len(rads) == 1
    assert re.search(r"(?:</m:radPr>|<m:rad>)<m:deg[ />]", flat(rads[0]))


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("part, tag", [("word/footnotes.xml", "footnote"), ("word/endnotes.xml", "endnote")])
def test_note_number_run_has_its_properties_first_with_separators_and_a_space(part, tag):
    """The number run of a note from add_footnote / add_endnote is w:rPr then the reference mark, the notes part
    holds the separator and continuation separator notes, and a space separates the number from the text. The w:rPr
    used to follow the mark, with no separator notes and no space (tensorbee/rdocx#432)."""
    doc = rdocx.Document()
    p = doc.add_paragraph("Body text")
    doc.add_footnote(p, "Source note")
    doc.add_endnote(p, "End note")
    xml = doc.part_xml(part)
    number = flat(re.search(rf"<w:r>(?:(?!</w:r>).)*<w:{tag}Ref/>.*?</w:r>", xml, re.S).group(0))
    assert "<w:rPr>" in number and number.index("<w:rPr>") < number.index(f"<w:{tag}Ref/>")
    assert len(re.findall(r'w:type="(?:separator|continuationSeparator)"', xml)) == 2
    assert "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml)).startswith(" ")


def with_hyperlink_style():
    doc = rdocx.Document()
    doc.add_style("Hyperlink", style_type="character", color="0563C1", underline=True)
    return doc


def paragraph_link():
    doc = with_hyperlink_style()
    doc.add_paragraph("See ").add_hyperlink("the site", "https://example.com")
    return doc


def heading_link():
    doc = with_hyperlink_style()
    doc.add_heading("Target", 1)
    doc.add_paragraph("See ")
    doc.add_heading_hyperlink(doc.paragraphs[1], doc.paragraphs[0], "Target")
    return doc


LINKS = {
    "Paragraph.add_hyperlink": paragraph_link,
    "Document.add_heading_hyperlink": heading_link,
    "Document.from_markdown": lambda: rdocx.Document.from_markdown(
        "See [the site](https://example.com).", template=with_hyperlink_style())[0],
    "Document.from_html": lambda: rdocx.Document.from_html('<p>See <a href="https://example.com">the site</a>.</p>')[0],
}


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("writer", LINKS)
def test_links_written_by_rdocx_take_the_hyperlink_style(writer):
    """The run of a link written by add_hyperlink, add_heading_hyperlink, from_markdown or from_html carries the
    Hyperlink character style (or at least a colour and an underline), so it looks like a link. The link run used to
    have no run properties (tensorbee/rdocx#432)."""
    doc = LINKS[writer]()
    link = flat(re.search(r"<w:hyperlink\b.*?</w:hyperlink>", doc.part_xml("word/document.xml"), re.S).group(0))
    assert '<w:rStyle w:val="Hyperlink"/>' in link or ("<w:color " in link and "<w:u " in link)


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("shape, first, second, expected", [
    ((1, 2), "Total", None, ["Total"]),
    ((1, 2), None, "Sum", ["Sum"]),
    ((1, 2), "Total", "Sum", ["Total", "Sum"]),
    ((2, 1), "Total", None, ["Total"]),
], ids=["horizontal-second-empty", "horizontal-first-empty", "horizontal-both-text", "vertical-second-empty"])
def test_merge_drops_the_empty_paragraph_of_an_empty_cell(shape, first, second, expected):
    """Cell.merge does not carry the lone empty paragraph of an empty cell into the merged cell (python-api.md,
    Tables, merge). Merging "Total" with an empty cell used to give ['Total', ''] (tensorbee/rdocx#432)."""
    rows, cols = shape
    other = (0, 1) if cols == 2 else (1, 0)
    doc = rdocx.Document()
    table = doc.add_table(rows, cols)
    if first:
        table.cell(0, 0).text = first
    if second:
        table.cell(*other).text = second
    doc.tables[0].cell(0, 0).merge(doc.tables[0].cell(*other))
    assert [p.text for p in doc.tables[0].cell(0, 0).paragraphs] == expected


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("how", ["grid_span", "merge"])
def test_a_cell_holding_only_paragraph_properties_merges_as_an_empty_cell(how):
    """An empty cell whose paragraph carries properties (a right-aligned totals column, or a cell made by
    insert_column next to one) merges as an empty cell: set_cell_grid_span and Cell.merge used to raise "horizontal
    merge cannot consume a nonempty cell" (python-api.md, Tables, merge)."""
    doc = rdocx.Document()
    table = doc.add_table(1, 3, style="Table Grid")
    table.cell(0, 0).text = "Total"
    table.cell(0, 1).paragraphs[0].alignment = rdocx.WD_ALIGN_PARAGRAPH.RIGHT
    if how == "grid_span":
        doc.tables[0].set_cell_grid_span(0, 0, 2)
    else:
        doc.tables[0].cell(0, 0).merge(doc.tables[0].cell(0, 1))
    assert [p.text for p in doc.tables[0].cell(0, 0).paragraphs] == ["Total"]


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("edit, cell", [("insert_column", (0, 2)), ("add_row", (2, 1))])
def test_new_cells_keep_the_paragraph_formatting_of_their_source(edit, cell):
    """A cell made by insert_column or add_row gets the source cell's paragraph properties and the run properties of
    its paragraph mark (python-api.md, Tables). New cells used to copy the shading only, losing the alignment and the
    paragraph-mark bold (tensorbee/rdocx#432)."""
    doc = rdocx.Document()
    doc.add_table(2, 2)
    for r in range(2):
        for c in range(2):
            source = doc.tables[0].cell(r, c)
            source.text = f"r{r}c{c}"
            source.shading = "D9E2F3"
            source.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            source.paragraphs[0].runs[0].font.bold = True
    if edit == "insert_column":
        doc.tables[0].insert_column(1, before=False)
    else:
        doc.tables[0].add_row()
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", doc.tables[0].cell(*cell).xml, re.S)
    assert ppr is not None
    ppr = flat(ppr.group(0))
    assert '<w:jc w:val="center"/>' in ppr
    assert re.search(r"<w:rPr>(?:(?!</w:rPr>).)*<w:b/>", ppr)


def add_comment(doc):
    doc.add_paragraph("text")
    doc.add_comment_on_text("text", author="A", text="a comment")


EDITS = {
    "no-other-edit": lambda doc: None,
    "add_footnote": lambda doc: doc.add_footnote(doc.add_paragraph("text"), "a note"),
    "add_picture": lambda doc: doc.add_picture(png_1x1(), "dot.png", width=rdocx.Inches(1), height=rdocx.Inches(1)),
    "add_comment_on_text": add_comment,
}


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("edit", EDITS)
def test_numbering_definition_id_stays_valid_after_an_edit(edit):
    """The id add_numbering_definition returns names its definition for the life of the document, whatever edit
    comes between (python-api.md, add_numbering_definition). After add_footnote or add_picture,
    add_numbering_instance(bullets) used to make an instance of the other, numbered definition
    (tensorbee/rdocx#432)."""
    doc = rdocx.Document()
    bullets = doc.add_numbering_definition([rdocx.ListLevel(format="bullet", text="o")])
    numbers = doc.add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.")])
    doc.add_paragraph("numbered item", style="List Paragraph").numbering = (doc.add_numbering_instance(numbers), 0)
    doc.add_paragraph("bullet item", style="List Paragraph").numbering = (doc.add_numbering_instance(bullets), 0)
    EDITS[edit](doc)
    instance = doc.add_numbering_instance(bullets)
    xml = doc.part_xml("word/numbering.xml")
    formats = dict(re.findall(r'<w:abstractNum w:abstractNumId="(\d+)".*?<w:numFmt w:val="(\w+)"', xml, re.S))
    abstract = dict(re.findall(r'<w:num w:numId="(\d+)"><w:abstractNumId w:val="(\d+)"/>', xml))
    assert formats[abstract[str(instance)]] == "bullet"


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
@pytest.mark.parametrize("toc_first", [False, True], ids=["toc-after-headings", "toc-before-headings"])
def test_rebuild_toc_fills_a_toc_inserted_before_the_headings(toc_first):
    """rebuild_toc() fills a TOC whatever the document held when insert_toc() wrote it (python-api.md, insert_toc may
    come before the headings). A TOC inserted before any heading used to make rebuild_toc() raise ("table of contents
    result must span paragraph boundaries", tensorbee/rdocx#432)."""
    doc = rdocx.Document()
    doc.add_paragraph("Contents")
    if toc_first:
        doc.insert_toc(1, max_level=3)
    doc.add_heading("Introduction", 1)
    doc.add_paragraph("Body text.")
    doc.add_heading("Methods", 1)
    if not toc_first:
        doc.insert_toc(1, max_level=3)
    doc.rebuild_toc()
    entries = re.findall(r'<w:hyperlink w:anchor="_Toc\w*"[^>]*><w:r><w:t>([^<]*)', doc.part_xml("word/document.xml"))
    assert entries == ["Introduction", "Methods"]


# carried by tensorbee/rdocx#441 (#432), delete once upstream main has it
def test_replace_agrees_with_text_on_chart_labels(tmp_path, rdocx_cli):
    """`rdocx replace` and try_replace_text search what `rdocx text` shows, so --expect takes its count, and a chart's
    cache and embedded workbook keep the same labels (python-api.md, try_replace_text). Both used to rewrite the
    chart cache labels, which `rdocx text` does not show, and leave the workbook with the old label
    (tensorbee/rdocx#432)."""
    doc = rdocx.Document()
    doc.add_paragraph("Sales in the North region rose.")
    data = rdocx.CategoryChartData()
    data.categories = ["North", "South"]
    data.add_series("Sales", [3, 4])
    doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, rdocx.Inches(4), rdocx.Inches(3))
    src = tmp_path / "chart.docx"
    doc.save(str(src))
    shown = run([rdocx_cli, "text", src]).stdout.count("North")
    res = run([rdocx_cli, "replace", src, "-p", "North", "-v", "NORTH", "--expect", shown, "-o", tmp_path / "out.docx"])
    assert res.returncode == 0, res.stderr
    edited = rdocx.Document(str(src))
    assert edited.try_replace_text("North", "NORTH") == shown
    package = zipfile.ZipFile(io.BytesIO(edited.to_bytes()))
    chart = next(n for n in package.namelist() if re.fullmatch(r"word/charts/chart\d+\.xml", n))
    cache = re.findall(r"<c:pt idx=\"\d\"><c:v>([A-Za-z]+)</c:v>", package.read(chart).decode())
    book = zipfile.ZipFile(io.BytesIO(package.read(next(n for n in package.namelist()
                                                        if n.startswith("word/embeddings/")))))
    in_book = "North" in "".join(book.read(n).decode() for n in book.namelist()
                                 if n == "xl/sharedStrings.xml" or n.startswith("xl/worksheets/"))
    assert ("North" in cache) == in_book



def toc_entry_looks(doc):
    """(fill, underlined) of the TOC entries Introduction and Methods in the SVG render of page 1."""
    svg = doc.render_page_to_svg(0).svg
    lines = [(float(x1), float(y)) for x1, y in re.findall(r'<line x1="([\d.]+)" y1="([\d.]+)"', svg)]
    looks = []
    for name in ("Introduction", "Methods"):
        m = re.search(rf'<text [^>]*y="([\d.]+)"[^>]*fill="(#[0-9A-Fa-f]{{6}})"[^>]*x="([\d.]+)[^"]*"[^>]*>{name}<',
                      svg)
        y, fill, x = float(m.group(1)), m.group(2).upper(), float(m.group(3))
        looks.append((fill, any(abs(lx - x) < 0.5 and 0 < ly - y < 3 for lx, ly in lines)))
    return looks


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
@pytest.mark.parametrize("word_runs", [False, True], ids=["rdocx-toc", "word-run-style"])
def test_toc_entries_are_drawn_without_the_hyperlink_look(word_runs):
    """The entries of a TOC are drawn in the TOC paragraph styles, black and not underlined, also when their runs
    carry w:rStyle "Hyperlink" as Word writes them. Such entries used to be drawn in the Hyperlink colour and
    underlined (tensorbee/rdocx#433)."""
    doc = rdocx.Document()
    doc.add_style("Hyperlink", style_type="character", color="0563C1", underline=True)
    doc.add_paragraph("Contents")
    for heading in ("Introduction", "Methods"):
        doc.add_heading(heading, 1)
        doc.add_paragraph("Body text.")
    doc.insert_toc(1, 3)
    doc.rebuild_toc()
    if word_runs:
        doc = patched(doc, {"word/document.xml": lambda x: re.sub(
            r'(<w:hyperlink w:anchor="_Toc[^"]*"[^>]*>)<w:r>', r'\1<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>',
            x)})
        assert '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Introduction' in doc.part_xml("word/document.xml")
    assert toc_entry_looks(doc) == [("#000000", False), ("#000000", False)]


INDENT_STYLES = (
    '<w:style w:type="paragraph" w:styleId="BaseInd"><w:name w:val="Base Ind"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="BaseInd"/>'
    '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="DerivedH"><w:name w:val="Derived H"/><w:basedOn w:val="BaseInd"/>'
    '<w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:style>')
INDENT_CASES = (
    '<w:pStyle w:val="BaseInd"/>',
    '<w:pStyle w:val="Derived"/>',
    '<w:pStyle w:val="BaseInd"/><w:ind w:left="0" w:firstLine="0"/>',
    '<w:pStyle w:val="DerivedH"/>',
)


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
def test_first_line_indent_is_one_property_across_style_levels():
    """w:firstLine and w:hanging are one property: the most specific level that sets either decides the first-line
    indent, so every first line here starts at the left margin (72 pt). A w:firstLine set by a derived style or by
    the paragraph used to keep the base style's hanging indent, starting the line at 36 pt (tensorbee/rdocx#433)."""
    doc = rdocx.Document()
    for k in range(len(INDENT_CASES)):
        doc.add_paragraph(f"Case{k} text")

    def body(xml):
        for k, ppr in enumerate(INDENT_CASES):
            xml = xml.replace(f"<w:p><w:r><w:t>Case{k} ", f"<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>Case{k} ")
        return xml

    doc = patched(doc, {"word/styles.xml": lambda x: x.replace("</w:styles>", INDENT_STYLES + "</w:styles>"),
                        "word/document.xml": body})
    svg = doc.render_page_to_svg(0).svg
    starts = [float(re.search(rf'<text [^>]*x="([\d.]+)[^"]*">Case{k}', svg).group(1))
              for k in range(len(INDENT_CASES))]
    assert starts == pytest.approx([72.0] * len(INDENT_CASES), abs=0.05)


def ruled_lines(png):
    """(horizontal, vertical) dark lines longer than 2 inches at 72 dpi; neighbouring pixel rows count once."""
    width, height, rows = decode_png(png)
    dark = [[max(px) < 200 for px in row] for row in rows]

    def runs(flags):
        return sum(1 for i, f in enumerate(flags) if f and (i == 0 or not flags[i - 1]))

    return (runs([sum(dark[y]) > 144 for y in range(height)]),
            runs([sum(dark[y][x] for y in range(height)) > 25 for x in range(width)]))


NO_BORDERS = "".join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                     for e in ("top", "left", "bottom", "right", "insideH", "insideV"))


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
@pytest.mark.parametrize("borders, expected", [("", (3, 3)), (NO_BORDERS, (0, 0))], ids=["style-only", "all-none"])
def test_table_borders_override_the_style_grid(borders, expected):
    """A table's own w:tblBorders override its style's: a TableGrid table whose borders are all none is drawn with no
    ruled line, one without keeps its 3 + 3 lines. The style's grid used to be drawn anyway (tensorbee/rdocx#433)."""
    doc = rdocx.Document()
    doc.add_table(2, 2).style = "TableGrid"
    for r in range(2):
        for c in range(2):
            doc.tables[0].cell(r, c).text = f"cell {r}{c}"
    if borders:
        xml = re.sub(r"(<w:tblPr>.*?)(</w:tblPr>)", r"\1<w:tblBorders>" + borders + r"</w:tblBorders>\2",
                     doc.tables[0].xml, count=1, flags=re.S)
        doc.tables[0].replace_xml(xml)
    assert ruled_lines(doc.render_page_to_png(0, 72)) == expected


@functools.cache
def row_lines_per_page(valign):
    """The lines of a 40-line one-cell row after 15 filler paragraphs, per page of the SVG render."""
    doc = rdocx.Document()
    for i in range(15):
        doc.add_paragraph(f"Filler paragraph {i + 1}.")
    doc.add_table(1, 1).style = "TableGrid"
    doc.tables[0].cell(0, 0).text = "Row line 1"
    for i in range(2, 41):
        doc.tables[0].cell(0, 0).add_paragraph(f"Row line {i}")
    if valign is not None:
        doc.tables[0].cell(0, 0).vertical_alignment = valign
    return [len(re.findall(r">Row(?: |<)", doc.render_page_to_svg(p).svg)) for p in range(doc.page_count())]


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
@pytest.mark.parametrize("valign", [WD_ALIGN_VERTICAL.TOP, WD_ALIGN_VERTICAL.CENTER, WD_ALIGN_VERTICAL.BOTTOM],
                         ids=["top", "center", "bottom"])
def test_cell_vertical_alignment_does_not_stop_a_row_from_splitting(valign):
    """w:vAlign does not change where a row breaks: the row starts on page 1 and breaks within one line of where it
    breaks without w:vAlign (python-api.md, Tables, vertical_alignment). A cell with any w:vAlign used to move the
    whole row to the next page (tensorbee/rdocx#433)."""
    base, got = row_lines_per_page(None), row_lines_per_page(valign)
    assert len(got) == len(base)
    assert got[0] > 0 and abs(got[0] - base[0]) <= 1


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
def test_automatic_text_colour_contrasts_with_the_cell_fill():
    """Text in the automatic colour is drawn white on a dark cell fill and black on a light one: black, navy, none.
    It used to be drawn black on every fill (tensorbee/rdocx#433)."""
    cases = (("Auto on black", "000000"), ("Auto on navy", "1F3864"), ("Auto on white", None))
    doc = rdocx.Document()
    doc.add_table(1, 3)
    for c, (text, fill) in enumerate(cases):
        cell = doc.tables[0].cell(0, c)
        cell.text = text
        if fill:
            cell.shading = fill
    svg = doc.render_page_to_svg(0).svg
    drawn = [re.search(rf'<text [^>]*fill="(#[0-9a-fA-F]{{6}})"[^>]*>{text.split()[-1]}<', svg).group(1).upper()
             for text, _ in cases]
    assert drawn == ["#FFFFFF", "#FFFFFF", "#000000"]


def symbol_bullets(bullets):
    """A document with one bullet paragraph per (font, code point) of bullets, each with its own single-level
    definition whose w:lvlText is the code point in that font, as Word writes a symbol-font bullet."""
    doc = rdocx.Document()
    for font, ch in bullets:
        doc.add_bullet_list_item(f"{font} bullet U+{ord(ch):04X}")
    abstract = "".join(
        f'<w:abstractNum w:abstractNumId="{900 + k}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0">'
        f'<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" '
        f'w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' for k, (font, ch) in enumerate(bullets))
    nums = "".join(f'<w:num w:numId="{900 + k}"><w:abstractNumId w:val="{900 + k}"/></w:num>'
                   for k in range(len(bullets)))

    def numbering(xml):
        xml = re.sub(r"(<w:numbering\b[^>]*>)", lambda m: m.group(1) + abstract, xml, count=1)
        return xml.replace("</w:numbering>", nums + "</w:numbering>")

    def body(xml):
        k = iter(range(len(bullets)))
        return re.sub(r'<w:numId w:val="\d+"/>', lambda m: f'<w:numId w:val="{900 + next(k)}"/>', xml)

    return patched(doc, {"word/numbering.xml": numbering, "word/document.xml": body})


def private_use(text):
    return any(0xE000 <= ord(c) <= 0xF8FF for c in text)


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
def test_wingdings_2_bullet_is_drawn_as_unicode():
    """A Wingdings 2 bullet square (U+F0A1 in w:lvlText) is drawn as a Unicode character, as the Wingdings (U+F0A7)
    and Symbol (U+F0B7) bullets are. It used to be drawn as its private-use code point, with no diagnostic
    (tensorbee/rdocx#433)."""
    bullets = (("Wingdings 2", "\uf0a1"), ("Wingdings", "\uf0a7"), ("Symbol", "\uf0b7"))
    svg = symbol_bullets(bullets).render_page_to_svg(0).svg
    drawn = re.findall(r'<text [^>]*x="90">([^<]+)<', svg)
    assert len(drawn) == len(bullets)
    assert [private_use(text) for text in drawn] == [False] * len(bullets)


# carried by tensorbee/rdocx#440 (#433), delete once upstream main has it
def test_unmapped_wingdings_2_code_is_reported():
    """A Wingdings 2 code point that has no Unicode mapping (U+F0A7 there) gets a render diagnostic naming it. Such a
    code used to be drawn as itself with no diagnostic (tensorbee/rdocx#433)."""
    result = symbol_bullets((("Wingdings 2", "\uf0a7"),)).render_page_to_svg(0)
    drawn = re.findall(r'<text [^>]*x="90">([^<]+)<', result.svg)
    reported = any("F0A7" in d.message.upper() and "Wingdings 2" in d.message for d in result.diagnostics)
    assert len(drawn) == 1
    assert not private_use(drawn[0]) or reported


DEJAVU_SANS = ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans.ttf",
               "/usr/share/fonts/TTF/DejaVuSans.ttf", "/Library/Fonts/DejaVuSans.ttf")


def pdf_page_count(pdf):
    return int(re.search(rb"/Type\s*/Pages\b.*?/Count (\d+)", pdf, re.S).group(1))


def pdf_fonts(pdf):
    return sorted({f.decode() for f in re.findall(rb"/BaseFont\s*/(?:[A-Z]{6}\+)?([A-Za-z0-9_-]+)", pdf)})


# upstream: tensorbee/rdocx#433
@pytest.mark.gap("docx-font-fallback-by-entry-point")
@pytest.mark.parametrize("folder", ["none", "empty", "one-font"])
def test_every_entry_point_resolves_a_font_the_same_way(folder, tmp_path, rdocx_cli, free_font):
    """For a font the document names and rdocx does not bundle (DejaVu Sans), `rdocx convert`, to_pdf(),
    page_count() and layout() give one font and one page count for the same font folder, and the CLI uses the font
    or reports the substitution. The CLI substitutes a bundled font silently, to_pdf() takes the system font, and
    page_count() / layout() follow one or the other depending on font_dir (tensorbee/rdocx#433)."""
    if not any(Path(p).is_file() for p in DEJAVU_SANS):
        pytest.skip("DejaVu Sans is not installed")
    doc = rdocx.Document()
    doc.set_style("Normal", font_name="DejaVu Sans")
    for _ in range(60):
        doc.add_paragraph("The committee reviews the annual budget, the schedule of works and the report of the "
                          "auditors for the coming year. " * 3)
    src = tmp_path / "fonts.docx"
    doc.save(str(src))
    fonts = tmp_path / "fonts"
    fonts.mkdir()
    if folder == "one-font":
        (fonts / free_font.name).write_bytes(free_font.read_bytes())
    kwargs, flags = ({}, []) if folder == "none" else ({"font_dir": str(fonts)}, ["--font-dir", fonts])
    res = run([rdocx_cli, "convert", src, "--to", "pdf", "-o", tmp_path / "cli.pdf", "--force", *flags])
    assert res.returncode == 0, res.stderr
    cli_pdf, py_pdf = (tmp_path / "cli.pdf").read_bytes(), doc.to_pdf(**kwargs)
    pages = {pdf_page_count(cli_pdf), pdf_page_count(py_pdf), doc.page_count(**kwargs),
             max(f.physical_page for f in doc.layout(**kwargs))}
    assert pdf_fonts(cli_pdf) == pdf_fonts(py_pdf)
    assert len(pages) == 1
    messages = [line for line in (res.stdout + res.stderr).splitlines() if not line.startswith("Written to")]
    assert pdf_uses_font(cli_pdf, b"DejaVuSans") or any("DejaVu Sans" in m for m in messages)


# ---- tensorbee/rdocx#439 (#435) --------------------------------------------------------------------------------------


def rezip(data, edits=None, duplicate=None):
    """The package data with the parts named in edits (name: function of the part's bytes) replaced and, when
    duplicate names a part, a second entry of that name, its "First copy" changed to "Second copy"."""
    edits = edits or {}
    zin, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            x = zin.read(item)
            z.writestr(item, edits[item.filename](x) if item.filename in edits else x)
            if item.filename == duplicate:
                with pytest.warns(UserWarning, match="Duplicate name"):
                    z.writestr(item.filename, x.replace(b"First copy", b"Second copy"))
    return out.getvalue()


def validate(cli, path, data):
    """Exit code and output of `rdocx validate` on data written to path."""
    path.write_bytes(data)
    res = run([cli, "validate", path])
    return res.returncode, res.stdout + res.stderr


def well_formed(xml):
    try:
        ET.fromstring(xml)
    except ET.ParseError:
        return False
    return True


RUN = "<w:r><w:t>Hello</w:t></w:r>"
RAD = ('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:rad><m:radPr>'
       '<m:degHide m:val="1"/></m:radPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>')
SCHEMA_FAULTS = {  # label: (old, new) on word/document.xml
    "w:pPr after w:r": (RUN, RUN + '<w:pPr><w:jc w:val="left"/></w:pPr>'),
    "w:rPr after w:t": (RUN, "<w:r><w:t>Hello</w:t><w:rPr><w:b/></w:rPr></w:r>"),
    "unknown element w:bogus": (RUN, "<w:bogus/>" + RUN),
    'w:jc w:val="middle"': (RUN, '<w:pPr><w:jc w:val="middle"/></w:pPr>' + RUN),
    'w:color w:val="#FF0000"': (RUN, '<w:r><w:rPr><w:color w:val="#FF0000"/></w:rPr><w:t>Hello</w:t></w:r>'),
    "text directly in w:r": (RUN, "<w:r>Hello</w:r>"),
    "w:sectPr not last in w:body": ("<w:sectPr>",
                                    "<w:sectPr></w:sectPr><w:p><w:r><w:t>after</w:t></w:r></w:p><w:sectPr>"),
    "w:ins without w:id": (RUN, f'<w:ins w:author="A">{RUN}</w:ins>'),
    "negative wp:extent": ('<wp:extent cx="914400"', '<wp:extent cx="-914400"'),
    "m:rad without m:deg": (RUN, RUN + RAD),
}


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_wordprocessingml_schema_faults(rdocx_cli, tmp_path):
    """`rdocx validate` exits non-zero naming word/document.xml on each of ten schema faults (child order, unknown
    element, enum values, text in w:r, w:sectPr not last, w:ins without w:id, negative extent, m:rad without m:deg),
    and passes the document they are made from (cli.md, validate). It used to exit 0 on all ten
    (tensorbee/rdocx#435)."""
    doc = rdocx.Document()
    doc.add_paragraph("Hello")
    doc.add_picture(png_1x1(), "dot.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
    doc.add_paragraph("World")
    base = doc.to_bytes()
    xml = zipfile.ZipFile(io.BytesIO(base)).read("word/document.xml").decode()
    path = tmp_path / "case.docx"
    assert validate(rdocx_cli, path, base)[0] == 0
    missed = {}
    for label, (old, new) in SCHEMA_FAULTS.items():
        assert old in xml, label
        bad = xml.replace(old, new, 1).encode()
        code, out = validate(rdocx_cli, path, rezip(base, {"word/document.xml": lambda x, bad=bad: bad}))
        if code == 0 or "document.xml" not in out:
            missed[label] = (code, out[:200])
    assert not missed


DANGLING = {  # label: (pattern, replacement, text the output must hold) on word/document.xml
    "a:blip r:embed names no relationship": (r'(<a:blip r:embed=")[^"]+"', r'\1rId999"', "rId999"),
    "w:headerReference r:id names no relationship": (r'(<w:headerReference [^>]*r:id=")[^"]+"', r'\1rId998"',
                                                     "rId998"),
    "w:footnoteReference to no footnote": (r'(<w:footnoteReference w:id=")\d+"', r'\g<1>99"', "99"),
    "w:t inside w:del (not w:delText)": (r"<w:delText([^>]*)>([^<]*)</w:delText>", r"<w:t\1>\2</w:t>",
                                         "document.xml"),
}


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_dangling_references_and_deleted_text_in_w_t(rdocx_cli, tmp_path):
    """`rdocx validate` exits non-zero on a picture or header reference to no relationship (naming the id), a
    footnote reference to no footnote (naming the id) and w:t inside w:del (naming the part), and passes the redline
    they are made from (cli.md, validate). It used to exit 0 on all four (tensorbee/rdocx#435)."""
    doc = rdocx.Document()
    doc.add_paragraph("Pay within thirty days.")
    edited = rdocx.Document()
    edited.add_paragraph("Pay within sixty days.")
    doc.compare(edited, "Reviewer", "2026-10-30T12:00:00Z", granularity="word")
    doc.set_header("Header text")
    doc.add_footnote(doc.add_paragraph("Footnoted."), "Note text")
    doc.add_picture(png_1x1(), "dot.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
    base = doc.to_bytes()
    xml = zipfile.ZipFile(io.BytesIO(base)).read("word/document.xml").decode()
    path = tmp_path / "case.docx"
    assert validate(rdocx_cli, path, base)[0] == 0
    missed = {}
    for label, (pattern, repl, needle) in DANGLING.items():
        bad = re.sub(pattern, repl, xml, count=1)
        assert bad != xml, label
        code, out = validate(rdocx_cli, path, rezip(base, {"word/document.xml": lambda x, bad=bad: bad.encode()}))
        if code == 0 or needle not in out:
            missed[label] = (code, out[:200])
    assert not missed


def unescaped_company(x):
    """docProps/app.xml with the Company property set to 'Smith & Sons', the ampersand left unescaped."""
    x = re.sub(rb"<Company>[^<]*</Company>|<Company/>", b"", x)
    return x.replace(b"</Properties>", b"<Company>Smith & Sons</Company></Properties>")


def first_copy_docx():
    doc = rdocx.Document()
    doc.add_paragraph("First copy")
    return doc.to_bytes()


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_a_malformed_properties_part_and_a_part_stored_twice(rdocx_cli, tmp_path):
    """`rdocx validate` exits non-zero naming the part on a docProps/app.xml that is not well-formed and on a ZIP
    that stores word/document.xml twice (cli.md, validate). It used to exit 0 on both (tensorbee/rdocx#435)."""
    path = tmp_path / "case.docx"
    code, out = validate(rdocx_cli, path, rezip(first_copy_docx(), {"docProps/app.xml": unescaped_company}))
    assert code != 0 and "app.xml" in out
    code, out = validate(rdocx_cli, path, rezip(first_copy_docx(), duplicate="word/document.xml"))
    assert code != 0 and "word/document.xml" in out


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_open_refuses_or_warns_on_a_part_stored_twice(tmp_path):
    """rdocx.Document() refuses, or warns about, a ZIP that stores word/document.xml twice, naming the part. It used
    to open it silently and read the last copy, where other readers read the first (tensorbee/rdocx#435)."""
    path = tmp_path / "case.docx"
    path.write_bytes(rezip(first_copy_docx(), duplicate="word/document.xml"))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            rdocx.Document(str(path))
            refusal = ""
        except Exception as e:
            refusal = str(e)
    assert "word/document.xml" in refusal or any("word/document.xml" in str(w.message) for w in caught)


def control_character_docx(path):
    """A document whose word/document.xml holds U+0001 inside its text, written to path."""
    doc = rdocx.Document()
    doc.add_paragraph("Hello")
    path.write_bytes(rezip(doc.to_bytes(), {"word/document.xml": lambda x: x.replace(b">Hello<", b">He\x01llo<")}))
    assert not well_formed(zipfile.ZipFile(path).read("word/document.xml"))
    return path


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_a_part_holding_a_control_character(rdocx_cli, tmp_path):
    """`rdocx validate` exits non-zero naming document.xml when it holds U+0001, which XML 1.0 forbids (cli.md,
    validate). It used to exit 0 (tensorbee/rdocx#435)."""
    path = control_character_docx(tmp_path / "ctrl.docx")
    res = run([rdocx_cli, "validate", path])
    assert res.returncode != 0 and "document.xml" in res.stdout + res.stderr


# upstream: tensorbee/rdocx#435
@pytest.mark.gap("docx-control-character-round-trip")
def test_open_refuses_a_control_character_or_save_writes_a_well_formed_part(tmp_path):
    """A document whose word/document.xml holds U+0001 is refused at open, naming the part, or opens and save()
    writes a well-formed part. Today it opens and save() writes the U+0001 back raw (tensorbee/rdocx#435)."""
    path = control_character_docx(tmp_path / "ctrl.docx")
    try:
        doc = rdocx.Document(str(path))
    except Exception as e:
        assert "document.xml" in str(e)
        return
    doc.save(str(tmp_path / "saved.docx"))
    assert well_formed(zipfile.ZipFile(tmp_path / "saved.docx").read("word/document.xml"))


BOOKMARK_CASES = {  # label: (part, text of the run to wrap, bookmark id, bookmark name) per bookmark
    "ids 1 and 2": [("document", "First heading", 1, "_Toc1"), ("document", "Some text", 2, "_GoBack")],
    "both id 2 in the body": [("document", "First heading", 2, "_Toc1"), ("document", "Some text", 2, "_GoBack")],
    "id 2 in the body and in the header": [("document", "First heading", 2, "_Toc1"),
                                           ("header", "Header text", 2, "_GoBack")],
}


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
@pytest.mark.parametrize("case", BOOKMARK_CASES)
def test_bookmarks_sharing_an_id_open_print_and_render(rdocx_cli, tmp_path, case):
    """A document in which two bookmarks share a w:id (Word's _GoBack reusing a _Toc id, in the body or in a header)
    opens, and `rdocx text` and `rdocx render` exit 0, as in Word and LibreOffice. All three used to refuse it
    ("duplicate bookmark id", tensorbee/rdocx#435)."""
    doc = rdocx.Document()
    doc.add_paragraph("First heading")
    doc.add_paragraph("Some text")
    doc.set_header("Header text")
    base = doc.to_bytes()
    parts = {"document": "word/document.xml",
             "header": next(n for n in zipfile.ZipFile(io.BytesIO(base)).namelist() if n.startswith("word/header"))}

    def bookmark(text, bid, name):
        def edit(x):
            wrapped = f"<w:r><w:t>{text}</w:t></w:r>".encode()
            assert wrapped in x, text
            return x.replace(wrapped, f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>'.encode() + wrapped
                             + f'<w:bookmarkEnd w:id="{bid}"/>'.encode(), 1)
        return edit

    data = base
    for part, text, bid, name in BOOKMARK_CASES[case]:
        data = rezip(data, {parts[part]: bookmark(text, bid, name)})
    path = tmp_path / "bookmarks.docx"
    path.write_bytes(data)
    rdocx.Document(str(path))
    assert run([rdocx_cli, "text", path]).returncode == 0
    assert run([rdocx_cli, "render", path, "-o", tmp_path / "pages", "--dpi", "20", "--force"]).returncode == 0


# ---- tensorbee/rdocx#442 (#436) --------------------------------------------------------------------------------------


RED = (200, 30, 30)


def colours(png):
    """(red pixels, pixels of another saturated colour) of a rendered PNG."""
    red = other = 0
    for row in decode_png(png)[2]:
        for r, g, b in row:
            if r > 150 and g < 80 and b < 80:
                red += 1
            elif max(r, g, b) - min(r, g, b) > 80:
                other += 1
    return red, other


def red_emf(w=200, h=100):
    """A minimal EMF: header, a solid red brush, one filled rectangle, end of file ([MS-EMF] 2.3)."""
    records = [
        struct.pack("<IIIIII", 39, 24, 1, 0, 0x001E1EC8, 0),  # EMR_CREATEBRUSHINDIRECT: handle 1, solid, red
        struct.pack("<III", 37, 12, 1),  # EMR_SELECTOBJECT: handle 1
        struct.pack("<II4i", 43, 24, 10, 10, w - 10, h - 10),  # EMR_RECTANGLE
        struct.pack("<IIIII", 14, 20, 0, 16, 20),  # EMR_EOF
    ]
    body = b"".join(records)
    header = struct.pack("<II4i4iIIIIHHIIIiiii", 1, 88, 0, 0, w - 1, h - 1, 0, 0, w * 2645 // 100, h * 2645 // 100,
                         0x464D4520, 0x10000, 88 + len(body), 1 + len(records), 2, 0, 0, 0, 0, 1024, 768, 270, 203)
    return header + body


def red_png(interlaced, w=64, h=48):
    """A uniformly red RGB PNG, plain or interlaced (Adam7: seven passes, each with its own filter bytes)."""
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    passes = ((0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2))
    raw = b"".join((b"\x00" + bytes(RED) * len(range(x0, w, dx))) * len(range(y0, h, dy))
                   for x0, y0, dx, dy in (passes if interlaced else ((0, 0, 1, 1),)) if x0 < w and y0 < h)
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 1 if interlaced else 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def category_data(categories, series, values, **kw):
    data = rdocx.CategoryChartData(**kw)
    data.categories = categories
    data.add_series(series, values)
    return data


def chart_document(data):
    doc = rdocx.Document()
    doc.add_paragraph("Chart")
    doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, rdocx.Inches(5), rdocx.Inches(3))
    return doc


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_render_draws_an_emf_picture():
    """An EMF picture (a red filled rectangle) is drawn in the page render. Its frame used to be left empty, with no
    diagnostic (tensorbee/rdocx#436)."""
    doc = rdocx.Document()
    doc.add_picture(red_emf(), "box.emf", width=rdocx.Inches(2), height=rdocx.Inches(1))
    png, _ = doc.render_page_to_png_with_diagnostics(0, 30)
    assert colours(png)[0] >= 1200  # two thirds of the 2 x 1 in frame at 30 dpi (1800 pixels)


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_render_draws_an_interlaced_png_like_its_plain_copy():
    """An interlaced (Adam7) PNG renders like its plain copy: the same red pixels, no other colour, no diagnostic. Its
    seven passes used to be read as the rows of a plain image and drawn as stripes of wrong colours
    (tensorbee/rdocx#436)."""
    def render(picture):
        doc = rdocx.Document()
        doc.add_picture(picture, "picture.png", width=rdocx.Inches(2), height=rdocx.Inches(1.5))
        return doc.render_page_to_png_with_diagnostics(0, 30)

    plain_red = colours(render(red_png(False))[0])[0]
    png, diagnostics = render(red_png(True))
    assert colours(png) == (plain_red, 0)
    assert not list(diagnostics)


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_a_chart_with_a_third_axis_id_is_read_and_drawn():
    """A column chart whose c:barChart lists a third c:axId naming no axis, as some generators write it, is read and
    drawn, the extra id ignored. Reading it used to raise ("ChartML contains duplicate c:axId") and the PDF showed a
    placeholder (tensorbee/rdocx#436)."""
    package = chart_document(category_data(["x", "y"], "A", (1, 2))).to_bytes()
    archive, out = zipfile.ZipFile(io.BytesIO(package)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in archive.infolist():
            data = archive.read(item)
            if re.fullmatch(r"word/charts/chart\d+\.xml", item.filename):
                data, n = re.subn(rb'(<c:axId val="\d+"/>)(</c:barChart>)', rb'\1<c:axId val="99999"/>\2', data)
                assert n == 1
            z.writestr(item, data)
    doc = rdocx.Document.from_bytes(out.getvalue())
    assert tuple(doc.charts[0].categories) == ("x", "y")
    assert not [str(d) for d in doc.to_pdf_with_diagnostics()[1] if "chart" in str(d)]


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
@pytest.mark.parametrize("fmt", ["General", "0", "0.00", "0.0%", "#,##0", "#,##0.0", "$#,##0", '"EUR "#,##0', "# ##0"])
def test_chart_number_formats_with_separators_currency_or_literals_draw_the_chart(fmt):
    """A chart number format with a thousands separator, a currency sign or a quoted literal draws the chart. Those
    used to turn the chart into an "Unsupported chart" placeholder ("ChartML c:numFmt/@formatCode projection has
    invalid value", tensorbee/rdocx#436)."""
    doc = chart_document(category_data(["Q1", "Q2"], "Revenue", (1820, 1910), number_format=fmt))
    diagnostics = [str(d) for d in doc.to_pdf_with_diagnostics()[1]]
    assert not [d for d in diagnostics if "formatCode" in d or "unsupported chart" in d]


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
@pytest.mark.parametrize("export", ["to_markdown", "to_html"])
def test_markdown_and_html_exports_hold_a_chart_as_a_table(export):
    """The Markdown and HTML exports write a chart as a table of its data: the category, the series name and the
    value are in the output. The chart used to be left out silently (tensorbee/rdocx#436)."""
    doc = rdocx.Document()
    doc.add_paragraph("Revenue")
    doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(["Q1", "Q2"], "Zeta region", (1150, 1190)),
                  rdocx.Inches(5), rdocx.Inches(3))
    text = getattr(doc, export)()
    assert [s for s in ("Q1", "Zeta region", "1150") if s not in text] == []


# ---- tensorbee/rdocx#444 (#434) --------------------------------------------------------------------------------------


C_STAMP = "2026-10-30T12:00:00Z"
C_SPLIT_TAIL = "ce is paid within thirty days."  # completes "Each invoi" into the original sentence


def commented_thirty():
    """A paragraph whose word "thirty" holds a comment by Ann, and the comment's id."""
    doc = rdocx.Document()
    doc.add_paragraph("Pay within thirty days.")
    root = doc.add_comment_on_text("thirty", author="Ann", initials="A", text="Why thirty?",
                                   date="2026-10-30T09:00:00Z")
    return doc, root


def mark_position(body, mark, cid):
    """Offset of <w:mark w:id="cid"/> in body, or None."""
    found = re.search(rf'<w:{mark} w:id="{cid}"/>', body)
    return found.start() if found else None


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_reply_to_writes_the_initials_of_the_reply():
    """reply_to takes initials, as add_comment does, and the reply in comments.xml carries them (python-api.md,
    reply_to). It used to raise TypeError on initials=, so a reply had no w:initials (tensorbee/rdocx#434)."""
    doc, root = commented_thirty()
    doc.reply_to(root, author="Bob", initials="B", text="Agreed")
    assert 'w:initials="B"' in doc.part_xml("word/comments.xml")


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_comment_reply_cli_takes_initials(tmp_path, rdocx_cli):
    """`rdocx comment reply` takes --initials (cli.md, comment reply). It used to fail with "unexpected argument
    '--initials' found" (tensorbee/rdocx#434)."""
    doc, root = commented_thirty()
    src = tmp_path / "comment.docx"
    doc.save(str(src))
    res = run([rdocx_cli, "comment", "reply", src, "--id", root, "--author", "Bob", "--initials", "B",
               "--text", "Agreed", "-o", tmp_path / "reply.docx"])
    assert res.returncode == 0, res.stderr


# upstream: tensorbee/rdocx#434
@pytest.mark.gap("docx-reply-without-range")
def test_reply_has_a_range_over_its_parent_text_and_a_reference_after_its_parent(tmp_path):
    """A reply added by reply_to gets its own w:commentRangeStart and w:commentRangeEnd covering its parent's text,
    and a w:commentReference after its parent's, as Word writes (python-api.md, reply_to). It gets none of the three
    in document.xml, so LibreOffice drops the reply on import (tensorbee/rdocx#434)."""
    doc, root = commented_thirty()
    reply = doc.reply_to(root, author="Bob", text="Agreed", date="2026-10-30T10:00:00Z")
    out = tmp_path / "reply.docx"
    doc.save(str(out))
    body = zipfile.ZipFile(out).read("word/document.xml").decode()
    marks = ("commentRangeStart", "commentRangeEnd", "commentReference")
    assert all(mark_position(body, m, root) is not None for m in marks)
    assert [m for m in marks if mark_position(body, m, reply) is not None] == list(marks)
    covered = re.sub(r"<[^>]+>", "", body[mark_position(body, "commentRangeStart", reply):
                                          mark_position(body, "commentRangeEnd", reply)])
    assert "thirty" in covered
    assert mark_position(body, "commentReference", reply) > mark_position(body, "commentReference", root)


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_comments_part_written_by_rdocx_holds_word_comment_markup():
    """comments.xml as rdocx creates it declares mc and lists w14 (used by w14:paraId) in mc:Ignorable, and each
    comment's first paragraph is in CommentText and opens with a CommentReference w:annotationRef run, both styles
    defined (python-api.md, comments). The root declared w14 without mc:Ignorable, and the comments had no
    annotationRef run and no comment styles (tensorbee/rdocx#434)."""
    doc, root = commented_thirty()
    doc.reply_to(root, author="Bob", text="Agreed")
    package = zipfile.ZipFile(io.BytesIO(doc.to_bytes()))
    xml = package.read("word/comments.xml").decode()
    styles = package.read("word/styles.xml").decode()
    start = re.search(r"<w:comments\b[^>]*>", xml).group(0)
    first_paras = [re.search(r"<w:p\b.*?</w:p>", c, re.S).group(0)
                   for c in re.findall(r"<w:comment\b.*?</w:comment>", xml, re.S)]
    assert "xmlns:mc=" in start
    ignorable = re.search(r'mc:Ignorable="([^"]*)"', start)
    assert ignorable is not None and "w14" in ignorable.group(1).split()
    annotation = re.compile(r'<w:pPr>.*?</w:pPr>\s*<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                            r'<w:annotationRef/></w:r>|<w:p\b[^>]*><w:r><w:rPr><w:rStyle w:val="CommentReference"/>'
                            r"</w:rPr><w:annotationRef/></w:r>", re.S)
    assert len(first_paras) == 2
    assert all(annotation.search(p) for p in first_paras)
    assert all('<w:pStyle w:val="CommentText"/>' in p for p in first_paras)
    assert 'w:styleId="CommentText"' in styles and 'w:styleId="CommentReference"' in styles


def split_invoice_redline(tail, granularity):
    """The redline of "Each invoice is paid within thirty days." against "Each invoi" + tail, the edited paragraph
    split inside "invoice" over two runs of one formatting: (revision count, deleted texts, inserted texts)."""
    original = rdocx.Document()
    original.add_paragraph("Each invoice is paid within thirty days.")
    edited = rdocx.Document()
    edited.add_paragraph("Each invoi").add_run(tail)
    out = rdocx.Document.from_bytes(original.to_bytes())
    out.compare(edited, "Reviewer", C_STAMP, granularity=granularity)
    xml = out.part_xml("word/document.xml")
    deleted = re.findall(r"<w:delText[^>]*>([^<]*)</w:delText>", xml)
    inserted = [re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<w:ins\b.*?</w:ins>", xml, re.S)]
    return len(out.revisions), deleted, inserted


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
@pytest.mark.parametrize("granularity", ["run", "word", "character"])
def test_compare_ignores_run_boundaries_inside_unchanged_text(granularity):
    """Two paragraphs with the same text and formatting compare equal whatever their run boundaries: no revision at
    any granularity (python-api.md, compare). At run and word granularity a word split over two runs used to be
    deleted and inserted again, 3 revisions (tensorbee/rdocx#434)."""
    assert split_invoice_redline(C_SPLIT_TAIL, granularity) == (0, [], [])


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_compare_by_word_over_split_runs_marks_only_the_edited_word():
    """thirty -> sixty in a paragraph split over runs of one formatting gives, at word granularity, one deletion
    ("thirty") and one insertion ("sixty") (python-api.md, compare). The split word "invoice" used to be deleted and
    inserted again as well (tensorbee/rdocx#434)."""
    _, deleted, inserted = split_invoice_redline("ce is paid within sixty days.", "word")
    assert (deleted, inserted) == (["thirty"], ["sixty"])


def late_fees_redline(path):
    """A word redline saved to path: "thirty" deleted, "sixty" inserted and "Late fees apply." made bold."""
    original = rdocx.Document()
    original.add_paragraph("Pay within thirty days.")
    original.add_paragraph("Late fees apply.")
    edited = rdocx.Document()
    edited.add_paragraph("Pay within sixty days.")
    edited.add_paragraph("Late fees apply.").runs[0].font.bold = True
    original.compare(edited, "Reviewer", C_STAMP, granularity="word")
    original.save(str(path))
    return original


C_REVISION_TEXT = {"deletion": "thirty", "insertion": "sixty"}  # any other kind: the property change


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_revisions_carry_their_text_in_python_and_in_the_cli_json(tmp_path, rdocx_cli):
    """Each revision, in Document.revisions and in `rdocx revision list --json`, has the text it inserts or deletes,
    or the text its property change covers (python-api.md, revisions; cli.md, revision list). Neither surface gave
    any text, so listing tracked changes with their content needed part_xml (tensorbee/rdocx#434)."""
    src = tmp_path / "redline.docx"
    doc = late_fees_redline(src)
    res = run([rdocx_cli, "revision", "list", src, "--json"])
    assert res.returncode == 0, res.stderr
    surfaces = {"python": [(r.kind, getattr(r, "text", None)) for r in doc.revisions],
                "cli": [(r["kind"], r.get("text")) for r in json.loads(res.stdout)["revisions"]]}
    for records in surfaces.values():
        assert len(records) == 3
        assert all(text == C_REVISION_TEXT.get(kind, "Late fees apply.") for kind, text in records)


def resolved_look(doc, text, prop):
    """Value of prop (font: w:rFonts/@w:ascii, size: w:sz/@w:val in half-points) for the paragraph holding text: its
    run, then its style chain through w:basedOn (the default paragraph style when it names none), then
    w:docDefaults."""
    attr = {"font": r'<w:rFonts [^>]*w:ascii="([^"]+)"', "size": r'<w:sz w:val="(\d+)"'}[prop]
    styles = doc.part_xml("word/styles.xml")
    para = re.search(rf"<w:p\b(?:(?!</w:p>).)*{re.escape(text)}.*?</w:p>", doc.part_xml("word/document.xml"),
                     re.S).group(0)
    first_run = re.search(r"<w:r\b[^>]*>\s*<w:rPr\b.*?</w:rPr>", para, re.S)
    if first_run and re.search(attr, first_run.group(0)):
        return re.search(attr, first_run.group(0)).group(1)
    named = re.search(r'<w:pStyle w:val="([^"]+)"/>', para)
    style_id = named.group(1) if named else next(
        re.search(r'w:styleId="([^"]+)"', t).group(1) for t in re.findall(r"<w:style [^>]*>", styles)
        if 'w:type="paragraph"' in t and 'w:default="1"' in t)
    while style_id:
        style = re.search(rf'<w:style [^>]*w:styleId="{style_id}".*?</w:style>', styles, re.S).group(0)
        rpr = re.search(r"<w:rPr\b.*?</w:rPr>", style, re.S)
        if rpr and re.search(attr, rpr.group(0)):
            return re.search(attr, rpr.group(0)).group(1)
        based = re.search(r'<w:basedOn w:val="([^"]+)"/>', style)
        style_id = based.group(1) if based else None
    found = re.search(r"<w:rPrDefault\b(?:(?!</w:rPrDefault>).)*?" + attr, styles, re.S)
    return found.group(1) if found else None


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
@pytest.mark.parametrize("source", ["docDefaults", "Normal, no w:pStyle", "Normal named by w:pStyle"])
def test_insert_document_keeps_the_font_and_size_of_the_inserted_document(source):
    """After insert_document, the inserted paragraph resolves to the font and size it has in the inserted document,
    Georgia set in its docDefaults or in its Normal style, named by the paragraph or not (python-api.md,
    insert_document). The docDefaults were dropped, and a paragraph naming no style fell on the target's Normal: the
    text took the target's Calibri 11 pt (tensorbee/rdocx#434)."""
    inserted = rdocx.Document()
    if source == "docDefaults":
        inserted.default_font_name = "Georgia"
        inserted.default_font_size = Pt(14)
    else:
        inserted.set_style("Normal", font_name="Georgia")
    inserted.add_paragraph("B body text.")
    if source == "Normal named by w:pStyle":
        inserted.paragraphs[0].style = "Normal"
    target = rdocx.Document()
    target.add_paragraph("A body text.")
    out = rdocx.Document.from_bytes(target.to_bytes())
    out.insert_document(inserted, at=1)
    for prop in ("font", "size"):
        assert resolved_look(out, "B body text.", prop) == resolved_look(inserted, "B body text.", prop)


def numbered_chapter(title):
    """A document holding a Heading 1 title and a numbered list of two steps."""
    doc = rdocx.Document()
    doc.add_heading(title, 1)
    for j in range(2):
        doc.add_numbered_list_item(f"{title}, step {j + 1}")
    return doc


# carried by tensorbee/rdocx#444 (#434), delete once upstream main has it
def test_insert_document_numbers_each_inserted_list_on_its_own():
    """With the default policy, each document inserted by insert_document keeps its list numbered on its own: three
    chapters numbered 1, 2 draw 1, 2, 1, 2, 1, 2, as with rename_all() (python-api.md, insert_document). An inserted
    list continued the target's equal list, drawn 1 to 6 (tensorbee/rdocx#434)."""
    doc = numbered_chapter("Chapter A")
    for title in ("Chapter B", "Chapter C"):
        doc.insert_document(numbered_chapter(title), at=len(doc.paragraphs) + len(doc.tables))
    assert re.findall(r">(\d+\.)<", doc.render_page_to_svg(0).svg) == ["1.", "2."] * 3


# ---- tensorbee/rdocx#443 (#437) --------------------------------------------------------------------------------------
# Timed as test_docx_scaling.py times its calls: the fastest of a few calls on each side, a generous ratio between the
# two sides and a floor, so that a loaded machine does not fail a flat call, where the cost the issue measured grew 9 to
# 56 times.
G_RATIO, G_FLOOR_MS = 3.0, 5


def fastest_ms(fn, n):
    """The fastest of n calls fn(i), in milliseconds."""
    best = float("inf")
    for i in range(n):
        t0 = time.perf_counter()
        fn(i)
        best = min(best, time.perf_counter() - t0)
    return 1000 * best


def with_styles(data, n):
    """The package `data` with n unused paragraph styles added to its styles part."""
    extra = "".join(
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Extra{i}"><w:name w:val="Extra {i}"/>'
        f'<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="{i}"/></w:pPr>'
        f'<w:rPr><w:sz w:val="{20 + i % 10}"/></w:rPr></w:style>' for i in range(n)).encode()
    zin, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            x = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                x = x.replace(b"</w:styles>", extra + b"</w:styles>")
            z.writestr(item, x)
    return out.getvalue()


def styles_document(styles):
    """A paragraph and a 3 x 6 table, with `styles` unused paragraph styles added (2000 make a 419 KB part)."""
    doc = rdocx.Document()
    doc.add_paragraph("Intro text")
    table = doc.add_table(3, 6)
    for r in range(3):
        for c in range(6):
            table.cell(r, c).text = f"r{r}c{c}"
    return rdocx.Document.from_bytes(with_styles(doc.to_bytes(), styles))


G_STYLE_EDITS = {
    "run.text": lambda d, i: setattr(d.paragraphs[0].runs[0], "text", f"Intro {i}"),
    "try_replace_text": lambda d, i: d.try_replace_text(f"r0c{i % 3}", f"r0c{i % 3}"),
    "add_row": lambda d, i: d.tables[0].add_row(),
    "clone_row": lambda d, i: d.tables[0].clone_row(1),
    "cell paragraph alignment": lambda d, i: setattr(d.tables[0].cell(0, 5).paragraphs[0], "alignment", i % 2),
}


# carried by tensorbee/rdocx#443 (#437), delete once upstream main has it
@pytest.mark.parametrize("edit", G_STYLE_EDITS)
def test_edit_cost_does_not_grow_with_the_styles_part(edit):
    """An edit costs about the same with rdocx's own 6 KB styles.xml and with 2000 unused styles (419 KB), as large
    as the styles part of python-docx's default template (gaps.md, edit costs). It used to cost 13 to 56 times more
    there, 30 to 125 ms a call (tensorbee/rdocx#437)."""
    op = G_STYLE_EDITS[edit]
    small, large = (fastest_ms(lambda i: op(d, i), 6) for d in (styles_document(0), styles_document(2000)))
    assert large < G_RATIO * small + G_FLOOR_MS, f"{edit}: {small:.1f} ms, then {large:.1f} ms with 2000 styles"


def cell_document(body):
    """A paragraph, a 40 x 6 table and `body` paragraphs after it, and a paragraph style TT."""
    doc = rdocx.Document()
    doc.add_paragraph("Intro text")
    table = doc.add_table(40, 6)
    for r in range(40):
        for c in range(6):
            table.cell(r, c).text = f"r{r}c{c}"
    for i in range(body):
        doc.add_paragraph(f"Body paragraph {i} with a few words.")
    doc.add_style("TT", based_on="Normal", font_size=rdocx.Pt(9))
    return doc


G_CELL_SETTERS = {
    "alignment": lambda p, i: setattr(p, "alignment", i % 2),
    "style": lambda p, i: setattr(p, "style", "TT" if i % 2 else "Normal"),
    "font.size": lambda p, i: setattr(p.runs[0].font, "size", rdocx.Pt(9 + i % 2)),
}


# carried by tensorbee/rdocx#443 (#437), delete once upstream main has it
@pytest.mark.parametrize("setter", G_CELL_SETTERS)
def test_cell_paragraph_setter_cost_does_not_grow_with_the_cell_place(setter):
    """A setter reached through a cell's paragraph or run costs about the same in the last row of a 40 x 6 table as in
    the first (python-api.md, Handles). It used to cost 10 to 17 times more there, 135 to 260 ms
    (tensorbee/rdocx#437)."""
    doc, op = cell_document(0), G_CELL_SETTERS[setter]
    first = fastest_ms(lambda i: op(doc.tables[0].cell(0, 5).paragraphs[0], i), 4)
    last = fastest_ms(lambda i: op(doc.tables[0].cell(39, 5).paragraphs[0], i), 4)
    assert last < G_RATIO * first + G_FLOOR_MS, f"{setter}: {first:.1f} ms in row 1, {last:.1f} ms in row 40"


# carried by tensorbee/rdocx#443 (#437), delete once upstream main has it
def test_cell_paragraph_setter_cost_does_not_grow_with_the_body():
    """A cell paragraph's setter costs about the same with 2000 body paragraphs after the table as with none, each
    timed after one untimed call on its document. It used to cost 8.6 times more (tensorbee/rdocx#437)."""
    def align(doc, i):
        doc.tables[0].cell(0, 5).paragraphs[0].alignment = i % 2

    costs = []
    for doc in (cell_document(0), cell_document(2000)):
        align(doc, 1)
        costs.append(fastest_ms(lambda i: align(doc, i), 4))
    alone, around = costs
    assert around < G_RATIO * alone + G_FLOOR_MS, f"{alone:.1f} ms alone, {around:.1f} ms with 2000 paragraphs"


def fill_per_call_ms(rows):
    """Milliseconds per call of setting a style on every cell paragraph of a fresh rows x 6 table, one call a cell."""
    doc = rdocx.Document()
    doc.add_paragraph("Intro text")
    doc.add_table(rows, 6)
    doc.add_style("TT", based_on="Normal", font_size=rdocx.Pt(9))
    cells = [(r, c) for r in range(rows) for c in range(6)]
    t0 = time.perf_counter()
    for r, c in cells:
        doc.tables[0].cell(r, c).paragraphs[0].style = "TT"
    return 1000 * (time.perf_counter() - t0) / len(cells)


# carried by tensorbee/rdocx#443 (#437), delete once upstream main has it
def test_filling_a_table_cell_by_cell_costs_the_same_per_cell():
    """Styling each cell paragraph of a 120-cell table costs about as much per call as in a 30-cell one, so a table
    fills in a time linear in its cells. It used to grow 4.5 times per call, the square of the cells for the table
    (tensorbee/rdocx#437)."""
    small = min(fill_per_call_ms(5) for _ in range(2))
    large = min(fill_per_call_ms(20) for _ in range(2))
    assert large < 2.5 * small + 2, f"{small:.1f} ms a call at 30 cells, {large:.1f} ms at 120"


# ---- tensorbee/rdocx#443 (#445) --------------------------------------------------------------------------------------
def body_and_table(body):
    """`body` paragraphs, then a fresh 4 x 5 table."""
    doc = rdocx.Document()
    for i in range(body):
        doc.add_paragraph(f"Body paragraph {i} with some text in it to look like a report.")
    doc.add_table(4, 5)
    return doc


G_CELL_PROPERTIES = {
    "width": lambda c, i: setattr(c, "width", rdocx.Inches(1 + i % 2)),
    "vertical_alignment": lambda c, i: setattr(c, "vertical_alignment", WD_ALIGN_VERTICAL.CENTER),
    "shading": lambda c, i: setattr(c, "shading", "D9E2F3"),
    "set_border": lambda c, i: c.set_border("top", "single", size=8, color="000000"),
    "set_margins": lambda c, i: c.set_margins(top=Pt(2), right=Pt(2), bottom=Pt(2), left=Pt(2)),
}


# carried by tensorbee/rdocx#443 (#445), delete once upstream main has it
@pytest.mark.parametrize("setter", G_CELL_PROPERTIES)
def test_cell_property_setter_cost_does_not_grow_with_the_body(setter):
    """A cell property setter edits the cell in place: on the fresh cells of a table after 2000 body paragraphs (the
    first property of a cell included) it costs about what a body paragraph's alignment costs (python-api.md,
    Handles). It used to cost 13 ms a call there, growing with the body (tensorbee/rdocx#445)."""
    doc, op = body_and_table(2000), G_CELL_PROPERTIES[setter]
    cells = [doc.tables[0].cell(r, c) for r in range(4) for c in range(5)]
    cell = fastest_ms(lambda i: op(cells[i], i), len(cells))
    paragraphs = doc.paragraphs
    paragraph = fastest_ms(lambda i: setattr(paragraphs[i], "alignment", i % 2), len(cells))
    assert cell < G_RATIO * paragraph + 2, f"{setter}: {cell:.2f} ms a cell, {paragraph:.2f} ms a paragraph"
    body = zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read("word/document.xml")
    assert body.count(b"<w:tcPr>") == len(cells), f"{setter}: not every cell got its property"


# carried by tensorbee/rdocx#443 (#445), delete once upstream main has it
@pytest.mark.parametrize("target", ["cell", "cell paragraph"])
def test_cell_text_costs_what_a_body_paragraph_text_costs(target):
    """`cell.text` and a cell paragraph's `text` cost about what a body paragraph's `text` costs with 2000 paragraphs
    in the body (gaps.md, edit costs). In the last row of the table they used to cost 3 and 7 times more
    (tensorbee/rdocx#445)."""
    doc = body_and_table(2000)
    if target == "cell":
        edit = lambda i: setattr(doc.tables[0].cell(3, i % 5), "text", f"v{i}")
    else:
        edit = lambda i: setattr(doc.tables[0].cell(3, i % 5).paragraphs[0], "text", f"v{i}")
    in_cell = fastest_ms(edit, 4)
    in_body = fastest_ms(lambda i: setattr(doc.paragraphs[i], "text", f"b{i}"), 4)
    assert in_cell < 2 * in_body + G_FLOOR_MS, f"{target}: {in_cell:.1f} ms, {in_body:.1f} ms for a body paragraph"
    assert rdocx.Document.from_bytes(doc.to_bytes()).tables[0].cell(3, 0).text == "v0"


# upstream: tensorbee/rdocx#445
@pytest.mark.gap("docx-python-replace-slower-than-cli")
def test_python_replacement_costs_what_the_cli_costs(tmp_path, rdocx_cli):
    """Document.open + try_replace_text + save cost at most 1.2 times `rdocx replace` on the same file of 4000
    paragraphs, process start included (gaps.md). They take about twice its time: the binding builds its owner
    inventory on open and again after the replacement (tensorbee/rdocx#445)."""
    doc = rdocx.Document()
    for i in range(4000):
        p = doc.add_paragraph()
        p.add_run(f"Section {i}: ").font.bold = True
        p.add_run("the plant treats " + ("water" if i % 4 == 0 else "air") + " for the town, ")
        p.add_run("as the report explains.").font.italic = True
    src = tmp_path / "in.docx"
    doc.save(str(src))
    python, cli = [], []
    for i in range(3):
        t0 = time.perf_counter()
        d = rdocx.Document.open(str(src))
        assert d.try_replace_text("water", "WATER") == 1000
        d.save(str(tmp_path / f"py{i}.docx"))
        python.append(time.perf_counter() - t0)
        t0 = time.perf_counter()
        res = run([rdocx_cli, "replace", src, "-p", "water", "-v", "WATER", "-o", tmp_path / f"cli{i}.docx"])
        cli.append(time.perf_counter() - t0)
        assert res.returncode == 0, res.stderr
    assert min(python) <= 1.2 * min(cli), f"Python {min(python):.2f} s, CLI {min(cli):.2f} s"
