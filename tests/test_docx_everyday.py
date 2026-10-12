"""docx: the everyday surface of the 2026-10-10 pin, written the python-docx way: authoring, lists, run and paragraph
formatting, styles, tables, pictures, links, headers and footers, notes, page setup, document-level views and
exports, properties, content controls, equations, raw XML, and the one-shot CLI commands; and from the 2026-10-12 pin,
documents from Markdown and HTML, charts and embedded fonts. Each test builds its own
input and checks the saved file (XML, python-docx, `rdocx validate`)."""
import datetime
import io
import json
import re
import zipfile

import docx
import pytest
import rdocx
from PIL import Image

from conftest import FREE_FONT_LICENSE, digest, libreoffice_pdf, part, pdf_uses_font, run

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def png(size=(60, 40), color=(200, 30, 30)):
    buf = io.BytesIO()
    Image.new("RGB", size, color).save(buf, "PNG")
    return buf.getvalue()


def saved(doc, path):
    doc.save(path)
    return part(path, "word/document.xml").decode()


# ---------------------------------------------------------------- authoring as python-docx does
def test_authoring_the_python_docx_way(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_heading("Report", 0)
    doc.add_heading("Scope", 1)
    p = doc.add_paragraph("Plain ")
    p.add_run("bold").bold = True
    doc.paragraphs[2].add_run(" italic").italic = True
    doc.paragraphs[2].runs[2].underline = True
    doc.add_paragraph("First point", style="List Bullet")
    doc.add_paragraph("Step one", style="List Number")
    doc.add_page_break()
    doc.add_table(2, 2, style="Table Grid")
    doc.paragraphs[2].insert_paragraph_before("Lead-in", style="Quote")
    assert doc.styles["Normal"].style_id == "Normal"
    buf = io.BytesIO()
    doc.save(buf)                                            # streams in and out
    back = rdocx.Document(io.BytesIO(buf.getvalue()))
    assert [p.text for p in back.paragraphs][:4] == ["Report", "Scope", "Lead-in", "Plain bold italic"]
    with pytest.raises(ValueError, match="to_pdf"):
        doc.save(tmp_path / "x.pdf")
    doc.save(tmp_path / "a.docx")
    d = docx.Document(tmp_path / "a.docx")
    assert [p.style.name for p in d.paragraphs[:6]] == ["Title", "Heading 1", "Quote", "Normal", "List Bullet", "List Number"]
    assert [(r.bold, r.italic, r.underline) for r in d.paragraphs[3].runs] == [(None, None, None), (True, None, None),
                                                                               (None, True, True)]
    assert d.tables[0].style.name == "Table Grid"
    assert '<w:br w:type="page"/>' in part(tmp_path / "a.docx", "word/document.xml").decode()
    assert run([rdocx_cli, "validate", tmp_path / "a.docx"]).returncode == 0


def test_lists_continue_restart_and_checklists(tmp_path):
    doc = rdocx.Document()
    doc.add_bullet_list_item("apple")
    doc.add_bullet_list_item("pear", level=1)
    doc.add_numbered_list_item("one")
    doc.add_numbered_list_item("two")
    doc.add_numbered_list_item("again one", restart=True)
    doc.add_numbered_list_item("again two")
    assert doc.paragraphs[0].numbering[1] == 0 and doc.paragraphs[1].numbering[1] == 1
    assert doc.paragraphs[2].numbering == doc.paragraphs[3].numbering
    assert doc.paragraphs[4].numbering[0] != doc.paragraphs[3].numbering[0]
    new_id = doc.restart_numbering(doc.paragraphs[5], start=5)
    assert doc.paragraphs[5].numbering == (new_id, 0)
    check = doc.add_numbering_definition([rdocx.ListLevel.checklist(checked=False)])
    plain = doc.add_numbering_definition([rdocx.ListLevel(format="bullet", text="-", font="Arial")])
    doc.add_paragraph("todo").numbering = (doc.add_numbering_instance(check), 0)
    doc.add_paragraph("dash").numbering = (doc.add_numbering_instance(plain, start=3), 0)
    doc.save(tmp_path / "l.docx")
    numbering = part(tmp_path / "l.docx", "word/numbering.xml").decode()
    assert '<w:startOverride w:val="5"/>' in numbering and "\u2610" in numbering
    assert re.search(r'<w:rFonts [^>]*w:ascii="Arial"', numbering)


def test_run_formatting_effects(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("x")
    for k in range(9):
        doc.paragraphs[0].add_run(str(k))
    runs = doc.paragraphs[0].runs
    runs[1].font.superscript = True
    runs[2].font.subscript = True
    runs[3].font.all_caps, runs[3].font.small_caps = True, True
    runs[4].font.double_strike = True
    runs[5].font.hidden = True
    runs[6].font.character_spacing = rdocx.Pt(2)
    runs[7].font.language, runs[7].font.east_asian_language, runs[7].font.complex_script_language = "fr-FR", "ja-JP", "ar-SA"
    runs[8].font.east_asian_name, runs[8].font.complex_script_name, runs[8].font.rtl = "MS Mincho", "Arial", True
    runs[9].font.highlight_color = rdocx.WD_COLOR_INDEX.YELLOW
    assert runs[1].font.superscript is True and runs[1].font.subscript is False
    runs[2].font.superscript = False                            # False clears only its own value, as python-docx
    assert runs[2].font.subscript is True
    assert runs[6].font.character_spacing == rdocx.Pt(2) and runs[9].font.highlight_color == rdocx.WD_COLOR_INDEX.YELLOW
    xml = saved(doc, tmp_path / "f.docx")
    for tag in ('<w:vertAlign w:val="superscript"/>', '<w:vertAlign w:val="subscript"/>', "<w:caps/>", "<w:smallCaps/>",
                "<w:dstrike/>", "<w:vanish/>", '<w:spacing w:val="40"/>', 'w:val="fr-FR"', 'w:eastAsia="ja-JP"',
                'w:bidi="ar-SA"', 'w:eastAsia="MS Mincho"', 'w:cs="Arial"', "<w:rtl/>", '<w:highlight w:val="yellow"/>'):
        assert tag in xml, tag
    r = docx.Document(tmp_path / "f.docx").paragraphs[0].runs
    assert (r[1].font.superscript, r[2].font.subscript, r[3].font.all_caps, r[4].font.double_strike) == (True, True, True, True)
    with pytest.raises(ValueError):
        doc.paragraphs[0].runs[7].font.language = "not a tag!"


def test_paragraph_format_tabs_borders_shading_outline_bidi(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Name\tValue")
    pf = doc.paragraphs[0].paragraph_format
    pf.tab_stops.add_tab_stop(rdocx.Inches(3), rdocx.WD_TAB_ALIGNMENT.RIGHT, rdocx.WD_TAB_LEADER.DOTS)
    pf.tab_stops.add_tab_stop(rdocx.Inches(1))
    assert [t.position for t in pf.tab_stops] == [rdocx.Inches(1), rdocx.Inches(3)]   # position order
    with pytest.raises(ValueError):
        pf.tab_stops.add_tab_stop(rdocx.Inches(1))              # one tab per position: change tab_stops[i]
    pf.tab_stops[0].alignment = rdocx.WD_TAB_ALIGNMENT.CENTER
    del pf.tab_stops[0]
    assert len(pf.tab_stops) == 1 and pf.tab_stops[0].leader == rdocx.WD_TAB_LEADER.DOTS
    pf.set_border("bottom", "single", size=8, color="#1F4E79")
    pf.set_border("top")
    assert pf.border("bottom") == ("single", 8, rdocx.RGBColor(0x1F, 0x4E, 0x79))
    pf.remove_border("top")
    pf.shading = "F2F2F2"
    pf.outline_level, pf.right_to_left = 1, True
    pf.line_spacing = 2                                          # a bare number is a multiple: double spacing
    xml = saved(doc, tmp_path / "p.docx")
    assert '<w:tab w:val="right" w:pos="4320" w:leader="dot"/>' in xml
    assert re.search(r'<w:pBdr><w:bottom w:val="single" w:sz="8"[^>]*w:color="1F4E79"', xml) and "<w:top " not in xml
    assert 'w:fill="F2F2F2"' in xml and '<w:outlineLvl w:val="1"/>' in xml and "<w:bidi/>" in xml
    assert docx.Document(tmp_path / "p.docx").paragraphs[0].paragraph_format.line_spacing == 2.0
    pf.clear_borders()
    pf.tab_stops.clear_all()
    assert pf.border("bottom") is None and len(pf.tab_stops) == 0


def test_style_keywords_and_document_defaults(tmp_path):
    doc = rdocx.Document()
    doc.add_style("Callout", based_on="Normal", underline=True, alignment=rdocx.WD_ALIGN_PARAGRAPH.CENTER,
                  line_spacing=1.5, keep_with_next=True, keep_together=True, page_break_before=False,
                  shading="#FFF2CC", borders={"left": ("single", 12, "C00000")},
                  tab_stops=[(rdocx.Inches(2), rdocx.WD_TAB_ALIGNMENT.DECIMAL)])
    doc.set_style("Callout", borders={"bottom": ("double", 4, (0, 0, 0))})   # replaces the whole set
    doc.default_font_name, doc.default_font_size = "Georgia", rdocx.Pt(11)
    assert (doc.default_font_name, doc.default_font_size) == ("Georgia", rdocx.Pt(11))
    doc.add_paragraph("boxed", style="Callout")
    doc.save(tmp_path / "s.docx")
    styles = part(tmp_path / "s.docx", "word/styles.xml").decode()
    callout = re.search(r'<w:style [^>]*w:styleId="Callout".*?</w:style>', styles, re.S).group(0)
    assert '<w:jc w:val="center"/>' in callout and "<w:keepNext/>" in callout and 'w:fill="FFF2CC"' in callout
    assert "<w:bottom " in callout and "<w:left " not in callout and 'w:val="decimal"' in callout
    assert re.search(r'<w:docDefaults>.*?w:ascii="Georgia".*?<w:sz w:val="22"/>', styles, re.S)
    with pytest.raises(ValueError):
        doc.add_style("Chars", "character", alignment=rdocx.WD_ALIGN_PARAGRAPH.CENTER)


def test_no_silent_failures(tmp_path):
    """A value that would silently do nothing raises and names the right call."""
    doc = rdocx.Document()
    doc.add_paragraph("x")
    with pytest.raises(ValueError, match="Pt"):
        doc.paragraphs[0].runs[0].font.size = 12                 # 12 EMU: a size in points needs Pt(12)
    with pytest.raises(ValueError, match="Pt"):
        doc.paragraphs[0].paragraph_format.space_after = 6
    with pytest.raises(ValueError, match="Twips"):
        doc.update_section(0, margin_top=1440)                     # twips passed as EMU
    with pytest.raises(AttributeError, match="update_section"):
        doc.sections[0].left_margin = rdocx.Inches(1)
    doc.update_section(0, margin_top=rdocx.Twips(1440), orientation="landscape")
    s = doc.sections[0]
    assert s.margin_top == rdocx.Inches(1) and s.orientation == "landscape" and s.page_width > s.page_height
    with pytest.raises(AttributeError, match="replace_xml|xml"):
        doc.paragraphs[0]._p


# ---------------------------------------------------------------- tables
def test_table_rows_columns_split_nested_and_look(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    t = doc.add_table(1, 2)
    cells = t.add_row().cells                                  # the python-docx way: handles stay valid
    cells[0].text, cells[1].text = "a", "b"
    col = t.add_column(rdocx.Inches(1))
    assert (col.index, col.width, len(col.cells)) == (2, rdocx.Inches(1), 2)
    t.insert_column(0, rdocx.Inches(0.5))
    t.remove_column(3)
    assert [c.index for c in t.columns] == [0, 1, 2] and t.cell(1, 1).text == "a"
    t.set_cell_grid_span(0, 1, 2)
    assert t.cell(0, 1).split() == 2                            # back to two cells
    inner = t.cell(1, 0).add_table(2, 2)
    inner.cell(0, 0).text = "nested"
    assert t.cell(1, 0).tables[0].cell(0, 0).text == "nested"
    t.first_row, t.horz_banding, t.first_col, t.last_row, t.last_col, t.vert_banding = True, True, False, False, False, False
    t.autofit = False
    t.indent = rdocx.Inches(0.2)
    with pytest.warns(UserWarning, match="default table style"):
        t.style = "No Such Style"
    t.style = "Table Grid"
    xml = saved(doc, tmp_path / "t.docx")
    assert re.search(r'<w:tblLook [^>]*w:firstRow="1"[^>]*w:noHBand="0"', xml) and '<w:tblLayout w:type="fixed"/>' in xml
    assert '<w:tblInd w:w="288" w:type="dxa"/>' in xml and xml.count("<w:tbl>") == 2
    d = docx.Document(tmp_path / "t.docx")
    assert len(d.tables[0].columns) == 3 and d.tables[0].cell(1, 0).tables[0].cell(0, 0).text == "nested"
    assert run([rdocx_cli, "validate", tmp_path / "t.docx"]).returncode == 0


# ---------------------------------------------------------------- pictures and links
def test_pictures_with_alt_text_wrap_crop_and_their_list(rdocx_cli, tmp_path):
    (tmp_path / "logo.png").write_bytes(png())
    doc = rdocx.Document()
    doc.add_picture(tmp_path / "logo.png", width=rdocx.Inches(1), description="Company logo", title="Logo")
    doc.add_picture(io.BytesIO(png()), height=rdocx.Inches(0.5), decorative=True, crop=(0.1, 0.0, 0.1, 0.0))
    doc.add_paragraph("Text that wraps around the picture.")
    pic = doc.paragraphs[-1].runs[0].add_picture(png(), rdocx.Inches(1), wrap="square",
                                                  position=("right", rdocx.Inches(0.2)), relative_to=("margin", "paragraph"))
    assert (pic.inline, pic.width) == (False, rdocx.Inches(1))
    doc.add_table(1, 1).cell(0, 0).paragraphs[0].add_run("in a cell").add_picture(png(), rdocx.Inches(0.3))
    assert len(doc.pictures) == 4
    pics = doc.pictures[:3]
    assert [(p.description, p.title, p.decorative, p.inline) for p in pics] == [
        ("Company logo", "Logo", False, True), (None, None, True, True), (None, None, False, False)]
    assert pics[0].content_type == "image/png" and pics[0].blob == png() and pics[0].height == rdocx.Inches(1) * 40 // 60
    assert doc.set_picture_size(pics[1], rdocx.Inches(2), rdocx.Inches(1)) == 1     # this picture only
    assert (doc.pictures[1].width, doc.pictures[0].width) == (rdocx.Inches(2), rdocx.Inches(1))
    with pytest.raises(ValueError):
        doc.add_picture(png(), wrap="sideways")
    xml = saved(doc, tmp_path / "p.docx")
    assert 'descr="Company logo"' in xml and 'title="Logo"' in xml and "decorative" in xml
    assert "<a:srcRect " in xml and "<wp:wrapSquare " in xml and '<wp:align>right</wp:align>' in xml
    assert run([rdocx_cli, "validate", tmp_path / "p.docx"]).returncode == 0


def test_internal_links_to_bookmarks_and_headings(tmp_path):
    doc = rdocx.Document()
    doc.add_heading("Results", 1)
    doc.add_paragraph("See ").add_hyperlink("the results", anchor=doc.paragraphs[0], tooltip="Jump")
    doc.add_paragraph("Web: ").add_hyperlink("site", "https://example.com")
    doc.add_paragraph("Again: ").add_hyperlink("results", "#" + doc.bookmarks[0].name)
    with pytest.raises(KeyError):
        doc.paragraphs[1].add_hyperlink("x", anchor="no_such_bookmark")
    links = doc.hyperlinks
    assert [(h.anchor, h.url, h.tooltip) for h in links] == [
        (doc.bookmarks[0].name, None, "Jump"), (None, "https://example.com", None), (doc.bookmarks[0].name, None, None)]
    xml = saved(doc, tmp_path / "l.docx")
    assert re.search(r'<w:hyperlink w:anchor="[^"]+" w:tooltip="Jump"', xml) or re.search(r'<w:hyperlink [^>]*w:anchor=', xml)
    assert b"#" not in part(tmp_path / "l.docx", "word/_rels/document.xml.rels")


# ---------------------------------------------------------------- headers, footers, notes, page setup
def test_headers_footers_page_numbers_the_python_docx_way(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Body")
    section = doc.sections[0]
    section.header.paragraphs[0].text = "Quarterly report"     # a new header holds one empty paragraph, as python-docx
    section.footer.add_page_number("Page {PAGE} of {NUMPAGES}")
    section.first_page_header.paragraphs[0].text = "Cover header"
    section.even_page_footer.add_paragraph("Even footer")
    t = section.header.add_table(1, 2, rdocx.Inches(6))
    t.cell(0, 1).text = "right cell"
    assert doc.sections[0].different_first_page_header_footer is True
    assert doc.settings.odd_and_even_pages_header_footer is True
    assert [p.text for p in doc.sections[0].header.paragraphs][0] == "Quarterly report"
    assert doc.sections[0].header.tables[0].rows[0].cells[1].text == "right cell"
    header_table = doc.sections[0].header.tables[0]
    header_table.cell(0, 0).add_paragraph("second line")
    assert (header_table.row_count, header_table.column_count) == (1, 2)
    assert [p.text for p in header_table.cell(0, 0).paragraphs] == ["", "second line"]
    assert (doc.sections[0].header.kind, doc.sections[0].header.variant, doc.sections[0].header.is_linked_to_previous) == (
        "header", "default", False)
    with pytest.raises(NotImplementedError):
        doc.sections[0].header.paragraphs[0].xml                           # raw XML: body and cells only
    doc.save(tmp_path / "h.docx")
    parts_ = {n: part(tmp_path / "h.docx", n).decode() for n in zipfile.ZipFile(tmp_path / "h.docx").namelist()
              if re.match(r"word/(header|footer)\d+\.xml$", n)}
    assert any('w:instr=" PAGE "' in x or "PAGE" in x for x in parts_.values())
    assert "<w:titlePg/>" in part(tmp_path / "h.docx", "word/document.xml").decode()
    assert "<w:evenAndOddHeaders/>" in part(tmp_path / "h.docx", "word/settings.xml").decode()
    d = docx.Document(tmp_path / "h.docx")
    assert d.sections[0].header.paragraphs[0].text == "Quarterly report"
    assert d.sections[0].first_page_header.paragraphs[0].text == "Cover header"
    assert "Page 1 of 1" in run([rdocx_cli, "text", tmp_path / "h.docx"], check=True).stdout


def test_a_new_section_unlinked_from_the_previous_header(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Portrait")
    doc.sections[0].header.paragraphs[0].text = "First section"
    doc.add_section(rdocx.WD_SECTION.NEW_PAGE)
    doc.update_section(1, orientation="landscape")
    doc.add_paragraph("Landscape")
    assert doc.sections[1].header.is_linked_to_previous is True
    doc.sections[1].header.is_linked_to_previous = False          # an empty header of its own
    doc.sections[1].header.paragraphs[0].text = "Second section"
    s0, s1 = doc.sections[0], doc.sections[1]
    assert (s1.orientation, s1.page_width, s1.margin_top) == ("landscape", s0.page_height, s0.margin_top)
    doc.save(tmp_path / "s.docx")
    d = docx.Document(tmp_path / "s.docx")
    assert [s.header.paragraphs[0].text for s in d.sections] == ["First section", "Second section"]
    assert d.sections[1].orientation == docx.enum.section.WD_ORIENT.LANDSCAPE


def test_footnotes_endnotes_and_their_numbering(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("A claim.")
    doc.add_paragraph("Another claim.")
    first = doc.add_footnote(doc.paragraphs[0], "The source.")
    doc.add_footnote(doc.paragraphs[1].runs[-1], "A second source.")
    doc.add_endnote(doc.paragraphs[1], "A closing note.")
    doc.remove_footnote(first)
    doc.set_note_numbering("endnote", number_format="lowerRoman")
    doc.save(tmp_path / "n.docx")
    text = run([rdocx_cli, "text", tmp_path / "n.docx"], check=True).stdout
    assert "A second source." in text and "A closing note." in text and "The source." not in text
    assert 'w:val="FootnoteReference"' in part(tmp_path / "n.docx", "word/document.xml").decode()
    assert '<w:numFmt w:val="lowerRoman"/>' in part(tmp_path / "n.docx", "word/settings.xml").decode()
    assert run([rdocx_cli, "validate", tmp_path / "n.docx"]).returncode == 0


def test_page_colour_watermarks_and_page_borders(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Draft body")
    doc.page_color = "#FDF6E3"
    assert doc.page_color == rdocx.RGBColor(0xFD, 0xF6, 0xE3)
    doc.set_text_watermark("DRAFT")
    doc.set_page_borders(0, "double", width=rdocx.Pt(1.5), color="1F4E79")
    xml = saved(doc, tmp_path / "w.docx")
    assert '<w:background w:color="FDF6E3"/>' in xml and re.search(r'<w:pgBorders [^>]*><w:top w:val="double"', xml)
    assert "<w:displayBackgroundShape/>" in part(tmp_path / "w.docx", "word/settings.xml").decode()
    headers = [n for n in zipfile.ZipFile(tmp_path / "w.docx").namelist() if n.startswith("word/header")]
    assert any("DRAFT" in part(tmp_path / "w.docx", n).decode() for n in headers)
    doc2 = rdocx.Document()
    doc2.set_image_watermark(png(), "mark.png", rdocx.Inches(3), rdocx.Inches(2))
    doc2.save(tmp_path / "i.docx")
    assert any(n.startswith("word/media/") for n in zipfile.ZipFile(tmp_path / "i.docx").namelist())
    assert run([rdocx_cli, "validate", tmp_path / "w.docx"]).returncode == 0


def test_a_picture_in_a_header_and_a_footer(rdocx_cli, tmp_path):
    """A logo next to the header text: the image is related from the header part (Word reports a picture related
    from document.xml as broken there), and doc.pictures stays the body's list."""
    doc = rdocx.Document()
    doc.add_paragraph("Body")
    doc.sections[0].header.paragraphs[0].text = "Logo: "
    logo = doc.sections[0].header.paragraphs[0].runs[0].add_picture(png(), rdocx.Inches(1), description="Logo")
    doc.sections[0].footer.paragraphs[0].text = "Seal "
    doc.sections[0].footer.paragraphs[0].runs[0].add_picture(png(), rdocx.Inches(0.5))
    assert (logo.content_type, logo.description, logo.width) == ("image/png", "Logo", rdocx.Inches(1))
    assert doc.pictures == ()
    with pytest.raises(ValueError, match="not in document.pictures"):
        doc.set_picture_size(logo, rdocx.Inches(2), rdocx.Inches(1))
    doc.save(tmp_path / "h.docx")
    names = zipfile.ZipFile(tmp_path / "h.docx").namelist()
    for story in ("header", "footer"):
        rels = [n for n in names if re.match(rf"word/_rels/{story}\d+\.xml\.rels$", n)]
        assert [n for n in rels if b"/image" in part(tmp_path / "h.docx", n)] != [], story
    assert b"/image" not in part(tmp_path / "h.docx", "word/_rels/document.xml.rels")
    assert run([rdocx_cli, "validate", tmp_path / "h.docx"]).returncode == 0
    header = docx.Document(tmp_path / "h.docx").sections[0].header
    blip = header._element.xpath(".//a:blip/@r:embed")[0]
    assert header.paragraphs[0].text == "Logo: " and header.part.related_parts[blip].blob == png()


# ---------------------------------------------------------------- document-level views, exports, checks
def test_views_exports_counts_and_validation(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_heading("Title words", 1)
    doc.add_paragraph("Three more words.")
    doc.sections[0].footer.paragraphs[0].text = "Footer text"
    doc.save(tmp_path / "v.docx")
    assert doc.text().rstrip("\n") == run([rdocx_cli, "text", tmp_path / "v.docx"], check=True).stdout.rstrip("\n")
    assert "Footer text" in doc.text()
    assert doc.to_markdown().startswith("# Title words") and "<h1>Title words</h1>" in doc.to_html()
    assert (doc.word_count(), doc.character_count(), doc.character_count(include_spaces=False)) == (5, 28, 25)
    assert doc.page_count() == 1
    with pytest.warns(rdocx.ConversionWarning, match="footer"):          # what the format cannot hold is named
        odt = doc.to_odt()
    with pytest.warns(rdocx.ConversionWarning):
        rtf = doc.to_rtf()
    with pytest.warns(rdocx.ConversionWarning):
        epub = doc.to_epub()
    assert zipfile.ZipFile(io.BytesIO(odt)).read("mimetype") == b"application/vnd.oasis.opendocument.text"
    assert rtf.startswith(b"{\\rtf1")
    assert zipfile.ZipFile(io.BytesIO(epub)).read("mimetype") == b"application/epub+zip"
    report = doc.validate()
    assert report.ok and report.errors == () and "Missing document title" in report.warnings
    with zipfile.ZipFile(tmp_path / "v.docx") as src, zipfile.ZipFile(tmp_path / "bad.docx", "w") as dst:
        for name in src.namelist():                          # a truncated footer part: the file no longer opens
            data = src.read(name)
            dst.writestr(name, data[:len(data) // 2] if re.match(r"word/footer\d+\.xml$", name) else data)
    bad = rdocx.Document.validate_file(tmp_path / "bad.docx")
    assert not bad.ok and any("footer" in e for e in bad.errors)
    assert run([rdocx_cli, "validate", tmp_path / "bad.docx"]).returncode == 1


def test_template_rendering_and_document_assembly(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Dear {{ client.name }},")
    doc.add_paragraph("{% for item in items %}")
    doc.add_paragraph("- {{ item }}")
    doc.add_paragraph("{% endfor %}")
    assert doc.render_template({"client": {"name": "Ada"}, "items": ["one", "two"]}) >= 3
    assert [p.text for p in doc.paragraphs] == ["Dear Ada,", "- one", "- two"]
    missing = rdocx.Document()
    missing.add_paragraph("{{ missing.value }}")
    with pytest.raises(rdocx.RdocxError, match="missing"):
        missing.render_template({})                             # names the tag; the document is unchanged
    assert missing.paragraphs[0].text == "{{ missing.value }}"
    annex = rdocx.Document()
    annex.add_style("Annex style", based_on="Normal")
    annex.add_paragraph("Annex body", style="Annex style")
    annex.add_paragraph("Annex end")
    doc.insert_document(annex)                                  # at the end, styles copied
    fragment = annex.copy_fragment(0, 1)                        # body items 0 up to 1 excluded
    doc.import_fragment(fragment, 0)
    assert [p.text for p in doc.paragraphs] == ["Annex body", "Dear Ada,", "- one", "- two", "Annex body", "Annex end"]
    assert doc.paragraphs[0].style == doc.paragraphs[4].style
    assert any(s.name == "Annex style" for s in doc.styles)
    doc.insert_document(annex, conflict="rename")               # an identical style is renamed instead of reused
    assert sum(s.name.startswith("Annex style") for s in doc.styles) == 2


def test_custom_app_and_settings_properties(tmp_path):
    doc = rdocx.Document()
    doc.custom_properties["Client"] = "Acme"
    doc.custom_properties["Revision"] = 3
    doc.custom_properties["Approved"] = True
    doc.custom_properties["Due"] = datetime.datetime(2026, 10, 31, tzinfo=datetime.timezone.utc)
    del doc.custom_properties["Approved"]
    doc.app_properties.company, doc.app_properties.manager = "Acme", "Ada"
    doc.settings.track_revisions = True
    doc.save(tmp_path / "c.docx")
    back = rdocx.Document.open(tmp_path / "c.docx")
    assert dict(back.custom_properties.items()) == {"Client": "Acme", "Revision": 3,
                                                    "Due": datetime.datetime(2026, 10, 31, tzinfo=datetime.timezone.utc)}
    assert (back.app_properties.company, back.app_properties.manager, back.settings.track_revisions) == ("Acme", "Ada", True)
    assert "Client" in back.custom_properties.keys() and back.custom_properties.get("Nope") is None
    assert "<w:trackRevisions/>" in part(tmp_path / "c.docx", "word/settings.xml").decode()


def controls_docx(path):
    """Body content controls: a plain text one (tag client), a drop-down (alias Status), a date and a check box."""
    d = docx.Document()
    body = d.element.body
    xml = (f'<w:p xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
           '<w:sdt><w:sdtPr><w:tag w:val="client"/><w:id w:val="11"/><w:showingPlcHdr/><w:text/></w:sdtPr>'
           '<w:sdtContent><w:r><w:t>Click here</w:t></w:r></w:sdtContent></w:sdt>'
           '<w:sdt><w:sdtPr><w:alias w:val="Status"/><w:tag w:val="status"/><w:id w:val="12"/><w:dropDownList>'
           '<w:listItem w:displayText="Draft" w:value="draft"/><w:listItem w:displayText="Final" w:value="final"/>'
           '</w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>Draft</w:t></w:r></w:sdtContent></w:sdt>'
           '<w:sdt><w:sdtPr><w:tag w:val="due"/><w:id w:val="13"/><w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>'
           '</w:sdtPr><w:sdtContent><w:r><w:t>date</w:t></w:r></w:sdtContent></w:sdt>'
           '<w:sdt><w:sdtPr><w:tag w:val="ok"/><w:id w:val="14"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>'
           '</w:sdtPr><w:sdtContent><w:r><w:t>\u2610</w:t></w:r></w:sdtContent></w:sdt></w:p>')
    from docx.oxml import parse_xml
    body.insert(0, parse_xml(xml))
    d.save(path)
    return path


def test_content_controls_are_listed_and_filled(rdocx_cli, tmp_path):
    doc = rdocx.Document.open(controls_docx(tmp_path / "c.docx"))
    assert [(c.tag, c.alias, c.id, c.text) for c in doc.content_controls] == [
        ("client", None, 11, "Click here"), ("status", "Status", 12, "Draft"), ("due", None, 13, "date"), ("ok", None, 14, "\u2610")]
    assert doc.set_content_control_value("Acme", tag="client") == 1
    doc.set_content_control_value("Final", alias="Status")
    doc.set_content_control_value("2026-10-31", tag="due")
    doc.set_content_control_value("true", tag="ok")
    with pytest.raises(ValueError, match="Draft|Final"):
        doc.set_content_control_value("Archived", alias="Status")       # a drop-down takes one of its items
    with pytest.raises(KeyError):
        doc.set_content_control_value("x", tag="nope")
    assert [c.text for c in doc.content_controls] == ["Acme", "Final", "2026-10-31", "\u2612"]
    assert [c.type for c in doc.content_controls] == ["plain_text", "dropdown_list", "date", "checkbox"]
    xml = saved(doc, tmp_path / "f.docx")
    assert "<w:showingPlcHdr/>" not in xml and 'w:fullDate="2026-10-31' in xml and '<w14:checked w14:val="1"/>' in xml
    res = run([rdocx_cli, "fill", tmp_path / "c.docx", "--tag", "client=Beta", "--alias", "Status=final", "-o", tmp_path / "g.docx", "--json"])
    assert res.returncode == 0, res.stderr
    assert [c.text for c in rdocx.Document.open(tmp_path / "g.docx").content_controls][:2] == ["Beta", "Final"]
    res = run([rdocx_cli, "fill", tmp_path / "c.docx", "--tag", "nope=1", "-o", tmp_path / "h.docx"])
    assert res.returncode == 1 and "client" in res.stderr                  # lists the known tags
    assert not (tmp_path / "h.docx").exists()


def test_equations_from_latex(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Energy: ").add_equation(r"E = mc^2")
    doc.add_paragraph("").add_equation(r"\frac{a}{b}", display=True)
    eqs = doc.paragraphs[0].equations
    assert len(eqs) == 1 and "mc" in eqs[0].latex and "<math" in eqs[0].mathml and eqs[0].display is False
    assert doc.paragraphs[1].equations[0].display is True
    with pytest.raises(ValueError):
        doc.paragraphs[0].add_equation(r"\frac{")
    assert "<m:oMath" in saved(doc, tmp_path / "e.docx")


# ---------------------------------------------------------------- raw XML
def test_raw_xml_of_paragraphs_runs_tables_cells_and_sections():
    doc = rdocx.Document()
    doc.add_paragraph("Plain text")
    doc.add_table(1, 1)
    p = doc.paragraphs[0]
    assert p.xml.startswith(b"<w:p") and b"Plain text" in p.xml
    p.replace_xml(p.xml.replace(b"<w:r>", b'<w:r><w:rPr><w:emboss/></w:rPr>', 1))   # a property with no API
    assert b"<w:emboss/>" in doc.paragraphs[0].runs[0].xml
    run_ = doc.paragraphs[0].runs[0]
    run_.replace_xml(run_.xml.replace(b"Plain", b"Raw"))
    assert doc.paragraphs[0].text == "Raw text"
    before = doc.to_bytes()
    for bad in (b"<w:p><w:bogus/></w:p>", b"<w:p>", b"<w:r/>", b'<w:p><w:hyperlink r:id="rId99"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>'):
        with pytest.raises(ValueError):
            doc.paragraphs[0].replace_xml(bad)
    assert doc.to_bytes() == before                                     # refused: nothing changed
    cell = doc.tables[0].cell(0, 0)
    assert cell.xml.startswith(b"<w:tc")
    cell.replace_xml(cell.xml.replace(b"</w:tc>", b"<w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc>"))
    assert [p.text for p in doc.tables[0].cell(0, 0).paragraphs] == ["", "second"]
    with pytest.raises(ValueError):
        doc.tables[0].cell(0, 0).replace_xml(b'<w:tc xmlns:w="%s"></w:tc>' % W.encode())   # a cell needs a paragraph
    table = doc.tables[0]
    table.replace_xml(table.xml)
    sect = doc.section_xml(0)
    assert sect.startswith(b"<w:sectPr")
    doc.replace_section_xml(0, sect.replace(b"<w:sectPr", b"<w:sectPr", 1).replace(b"</w:sectPr>", b"<w:lnNumType w:countBy=\"1\"/></w:sectPr>"))
    assert b"lnNumType" in doc.section_xml(0)


# ---------------------------------------------------------------- one-shot CLI commands
def test_cli_replace_map_and_regex(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Dear {{name}}, see you on 2026-10-31 at {{place}}.")
    doc.add_paragraph("{{name}} again.")
    doc.save(tmp_path / "t.docx")
    (tmp_path / "pairs.json").write_text(json.dumps([{"placeholder": "{{name}}", "value": "Ada", "expect": 2},
                                                     {"placeholder": "{{place}}", "value": "the office"}]))
    res = run([rdocx_cli, "replace", tmp_path / "t.docx", "--map", tmp_path / "pairs.json", "-o", tmp_path / "a.docx", "--json"])
    assert res.returncode == 0, res.stderr
    record = json.loads(res.stdout)
    assert ([p["count"] for p in record["pairs"]], record["total"]) == ([2, 1], 3)
    assert rdocx.Document.open(tmp_path / "a.docx").paragraphs[0].text == "Dear Ada, see you on 2026-10-31 at the office."
    (tmp_path / "bad.json").write_text(json.dumps([{"placeholder": "{{name}}", "value": "Ada", "expect": 3}]))
    res = run([rdocx_cli, "replace", tmp_path / "t.docx", "--map", tmp_path / "bad.json", "-o", tmp_path / "b.docx"])
    assert res.returncode == 1 and '"{{name}}"' in res.stderr and not (tmp_path / "b.docx").exists()  # names the pair
    res = run([rdocx_cli, "replace", tmp_path / "a.docx", "--regex", "-p", r"(\d{4})-(\d{2})-(\d{2})", "-v", "$3/$2/$1",
               "--expect", "1", "-o", tmp_path / "c.docx"])
    assert res.returncode == 0, res.stderr
    assert "31/10/2026" in rdocx.Document.open(tmp_path / "c.docx").paragraphs[0].text


def test_cli_fields_update_images_extract_meta_and_inspect(rdocx_cli, tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Printed ")
    doc.paragraphs[0].runs[0].add_field('DATE \\@ "yyyy-MM-dd"', "old")
    doc.add_picture(png(), "logo.png", rdocx.Inches(1), description="Logo")
    doc.add_picture(png(color=(0, 0, 200)), "chart.png", rdocx.Inches(1))
    doc.save(tmp_path / "d.docx")
    res = run([rdocx_cli, "fields", "update", tmp_path / "d.docx", "--now", "2026-10-10", "-o", tmp_path / "f.docx", "--json"])
    assert res.returncode == 0, res.stderr
    assert "2026-10-10" in rdocx.Document.open(tmp_path / "f.docx").paragraphs[0].text
    res = run([rdocx_cli, "images", "extract", tmp_path / "d.docx", tmp_path / "pics", "--json"])
    assert res.returncode == 0, res.stderr
    listing = json.loads(res.stdout)
    assert sorted(p.name for p in (tmp_path / "pics").iterdir()) == ["image1.png", "image2.png"]
    assert [(i["alt_text"], i["relationship_id"], i["format"]) for i in listing["images"]] == [("Logo", "rId1", "png"),
                                                                                               (None, "rId2", "png")]
    assert run([rdocx_cli, "images", "extract", tmp_path / "d.docx", tmp_path / "pics"]).returncode == 1     # existing files
    assert run([rdocx_cli, "images", "extract", tmp_path / "d.docx", tmp_path / "pics", "--force"]).returncode == 0
    res = run([rdocx_cli, "meta", "set", tmp_path / "d.docx", "--title", "Final report", "--author", "Ada", "--keywords", "q3",
               "--custom", "Client=Acme", "--custom", "Ref=42", "-o", tmp_path / "m.docx"])
    assert res.returncode == 0, res.stderr
    res = run([rdocx_cli, "meta", "set", tmp_path / "m.docx", "--remove-custom", "Ref", "--subject", "S", "--description", "D",
               "--category", "C", "-o", tmp_path / "m2.docx"])
    assert res.returncode == 0, res.stderr
    meta = json.loads(run([rdocx_cli, "meta", "get", "--json", tmp_path / "m2.docx"], check=True).stdout)
    text = json.dumps(meta)
    assert "Final report" in text and "Ada" in text and "Acme" in text and "Ref" not in text
    back = rdocx.Document.open(tmp_path / "m2.docx")
    assert (back.core_properties.title, back.custom_properties["Client"]) == ("Final report", "Acme")
    back.custom_properties["Count"] = 3
    back.save(tmp_path / "typed.docx")
    run([rdocx_cli, "meta", "set", tmp_path / "typed.docx", "--custom", "Count=4", "-o", tmp_path / "t2.docx"], check=True)
    assert rdocx.Document.open(tmp_path / "t2.docx").custom_properties["Count"] == 4      # an int stays an int
    assert run([rdocx_cli, "meta", "set", tmp_path / "typed.docx", "--custom", "Count=many", "-o", tmp_path / "t3.docx"]).returncode == 1
    info = json.loads(run([rdocx_cli, "inspect", "--json", tmp_path / "d.docx"], check=True).stdout)
    assert (info["words"], info["characters"], info["characters_no_spaces"], info["pages"]) == (1, 8, 7, 1)
    assert [p["alt_text"] for p in info["pictures"]] == ["Logo", None] and info["content_controls"] == []


def test_breaks_header_styles_and_nested_table_limits(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("Line one")
    doc.paragraphs[0].runs[0].add_break()                                  # a line break by default
    doc.paragraphs[0].add_run("Line two").add_break(rdocx.WD_BREAK.PAGE)
    doc.paragraphs[0].add_run("Column").add_break(rdocx.WD_BREAK.COLUMN)
    doc.sections[0].header.paragraphs[0].text = "Head"
    with pytest.raises((ValueError, rdocx.RdocxError)):
        doc.sections[0].header.paragraphs[0].runs[0].add_break(rdocx.WD_BREAK.PAGE)   # Word ignores it there
    inner = doc.add_table(1, 1).cell(0, 0).add_table(1, 1)
    inner.cell(0, 0).text = "deep"
    with pytest.raises(NotImplementedError):
        doc.tables[0].cell(0, 0).tables[0].cell(0, 0).paragraphs
    with pytest.raises(NotImplementedError):
        doc.tables[0].cell(0, 0).tables[0].xml
    xml = saved(doc, tmp_path / "b.docx")
    assert "<w:br/>" in xml and '<w:br w:type="page"/>' in xml and '<w:br w:type="column"/>' in xml
    docx.Document().save(tmp_path / "styled.docx")                # python-docx's template defines Header and Footer
    styled = rdocx.Document.open(tmp_path / "styled.docx")
    styled.sections[0].footer.add_paragraph("Foot")                  # takes the Footer style by itself
    styled.save(tmp_path / "s.docx")
    assert docx.Document(tmp_path / "s.docx").sections[0].footer.paragraphs[-1].style.name == "Footer"


def test_raw_xml_refusals_leave_the_document_unchanged():
    doc = rdocx.Document()
    doc.add_paragraph("Keep me")
    before = doc.to_bytes()
    for bad in (b'<!DOCTYPE w:p []><w:p/>',
                b'<w:p><x:foo xmlns:x="urn:example"/></w:p>',                                  # foreign, not ignorable
                b'<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>',  # a section break
                b'<w:p/><w:p/>'):
        with pytest.raises(ValueError):
            doc.paragraphs[0].replace_xml(bad)
    assert doc.to_bytes() == before


def test_a_heading_link_adds_a_heading_bookmark():
    doc = rdocx.Document()
    doc.add_heading("Results", 1)
    doc.add_paragraph("See ").add_hyperlink("results", anchor=doc.paragraphs[0])
    assert doc.bookmarks[0].name.startswith("Heading_") and doc.hyperlinks[0].anchor == doc.bookmarks[0].name


# ---------------------------------------------------------------- the 2026-10-12 pin: import, charts, embedded fonts
def compat_mode(path):
    return re.search(r'w:name="compatibilityMode"[^>]*w:val="(\d+)"', part(path, "word/settings.xml").decode())


def test_new_documents_name_compatibility_mode_15(tmp_path):
    """Word opens a document that names no compatibility mode in Compatibility Mode; new ones name 15."""
    rdocx.Document().save(tmp_path / "new.docx")
    rdocx.Document.from_markdown("# Title\n")[0].save(tmp_path / "md.docx")
    rdocx.Document.from_html("<p>x</p>")[0].save(tmp_path / "html.docx")
    for name in ("new.docx", "md.docx", "html.docx"):
        assert compat_mode(tmp_path / name).group(1) == "15", name
    with zipfile.ZipFile(tmp_path / "new.docx") as src, zipfile.ZipFile(tmp_path / "legacy.docx", "w") as dst:
        for info in src.infolist():                                       # a file that names no mode keeps none
            data = src.read(info.filename)
            if info.filename == "word/settings.xml":
                data = re.sub(rb"<w:compat>.*?</w:compat>", b"", data)
            dst.writestr(info, data)
    rdocx.Document.open(tmp_path / "legacy.docx").save(tmp_path / "resaved.docx")
    assert compat_mode(tmp_path / "resaved.docx") is None


def test_markdown_becomes_a_styled_document(rdocx_cli, tmp_path):
    (tmp_path / "pic.png").write_bytes(png())
    (tmp_path / "notes.md").write_text(
        "# Site visit\n\nVisited on **27 September**, see [the survey](https://example.com/s).\n\n"
        "- bearings\n  - pier 2\n- joints\n\n1. first\n2. second\n\n| Item | State |\n|---|---|\n| Pier 2 | Good |\n\n"
        "> A quote.\n\n```\ncode line\n```\n\n![Deck photo](pic.png)\n")
    template = rdocx.Document()
    template.sections[0].header.paragraphs[0].text = "Template header"
    template.save(tmp_path / "brand.docx")
    res = run([rdocx_cli, "convert", tmp_path / "notes.md", "--template", tmp_path / "brand.docx", "-o", tmp_path / "n.docx"])
    assert res.returncode == 0, res.stderr
    doc = rdocx.Document.open(tmp_path / "n.docx")
    styles = {p.text: p.style for p in doc.paragraphs}
    assert styles["Site visit"] == "Heading1" and styles["A quote."] == "Quote"
    assert doc.styles[styles["code line"]].name == "HTML Preformatted"
    assert [p.numbering is not None for p in doc.paragraphs if p.text in ("bearings", "pier 2", "first")] == [True] * 3
    levels = {p.text: p.numbering[1] for p in doc.paragraphs if p.numbering}
    assert (levels["bearings"], levels["pier 2"]) == (0, 1)
    assert doc.tables[0].cell(1, 0).text == "Pier 2" and doc.tables[0].rows[0].is_header
    assert [(h.text, h.url) for h in doc.hyperlinks] == [("the survey", "https://example.com/s")]
    assert any(r.text.strip() == "27 September" and r.bold for p in doc.paragraphs for r in p.runs)
    assert [p.description for p in doc.pictures] == ["Deck photo"]
    assert doc.sections[0].header.paragraphs[0].text == "Template header"
    assert run([rdocx_cli, "validate", tmp_path / "n.docx"]).returncode == 0
    # the default output is the input name with .docx; a DOCX input still needs --to
    assert run([rdocx_cli, "convert", tmp_path / "notes.md"]).returncode == 0 and (tmp_path / "notes.docx").exists()
    assert run([rdocx_cli, "convert", tmp_path / "n.docx", "-o", tmp_path / "x.pdf"]).returncode != 0
    # Python: the same import, with the diagnostics of what was dropped, and the folder images come from
    md, diagnostics = rdocx.Document.open_markdown(tmp_path / "notes.md")
    assert [p.description for p in md.pictures] == ["Deck photo"] and diagnostics == ()
    _, diagnostics = rdocx.Document.from_markdown("![missing](nowhere.png)\n", base_dir=tmp_path)
    assert diagnostics and diagnostics[0].location == "line 1" and diagnostics[0].message
    (tmp_path / "bad.md").write_text("![missing](nowhere.png)\n")
    res = run([rdocx_cli, "convert", tmp_path / "bad.md", "-o", tmp_path / "bad.docx"])
    assert res.returncode == 0 and "Warning: line 1:" in res.stderr


def test_html_becomes_a_document_and_inserts_after_an_item(rdocx_cli, tmp_path):
    pic = "data:image/png;base64," + __import__("base64").b64encode(png()).decode()
    html = (f"<h1>Findings</h1><p>See <a href='https://example.com/s'>the survey</a> and "
            f"<a href='javascript:alert(1)'>this</a>.</p><ul><li>Pier 2</li></ul><img src='{pic}' alt='Deck'>")
    doc, diagnostics = rdocx.Document.from_html(html)
    assert doc.paragraphs[0].style == "Heading1" and doc.paragraphs[0].text == "Findings"
    assert [h.url for h in doc.hyperlinks] == ["https://example.com/s"]          # the unsafe link keeps its text only
    assert "this" in doc.paragraphs[1].text and [p.description for p in doc.pictures] == ["Deck"]
    assert diagnostics and all(d.location and d.message for d in diagnostics)
    assert all(d.property is None or isinstance(d.property, str) for d in diagnostics)
    heading = next(item for item in doc.story_items if item.text == "Findings")
    assert doc.insert_html("<p><em>After the heading.</em></p>", after=heading) == ()
    assert doc.paragraphs[1].text == "After the heading." and doc.paragraphs[1].runs[0].italic
    (tmp_path / "page.html").write_text(html)
    res = run([rdocx_cli, "convert", tmp_path / "page.html", "-o", tmp_path / "page.docx"])
    assert res.returncode == 0 and "Warning:" in res.stderr
    opened, _ = rdocx.Document.open_html(tmp_path / "page.html")
    assert [p.text for p in opened.paragraphs] == [p.text for p in rdocx.Document.open(tmp_path / "page.docx").paragraphs]


def test_markdown_export_reads_as_the_document_means():
    """Edge spaces go outside the emphasis, two bold runs share one pair, a list ends before what follows."""
    doc = rdocx.Document()
    p = doc.add_paragraph("A")
    p.add_run(" rose").bold = True
    p.add_run(" bush").bold = True
    doc.add_bullet_list_item("item")
    doc.add_paragraph("After the list.")
    md = doc.to_markdown()
    assert "A **rose bush**" in md and "****" not in md and "** rose" not in md
    assert "- item\n\nAfter the list." in md


def test_charts_are_added_and_read_back(rdocx_cli, tmp_path):
    from rdocx.chart.data import CategoryChartData, XyChartData
    from rdocx.enum.chart import XL_CHART_TYPE
    doc = rdocx.Document()
    doc.add_heading("Costs", 1)
    doc.add_paragraph("End.")
    data = CategoryChartData()
    data.categories = ["2024", "2025"]
    data.add_series("Inspections", (12, 15))
    heading = next(item for item in doc.story_items if item.text == "Costs")
    item = doc.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, rdocx.Inches(5), rdocx.Inches(3), data, after=heading)
    assert item.direct_body_index == 1
    xy = XyChartData()
    s = xy.add_series("Deflection")
    for x, y in ((0, 1.5), (1, 2.5)):
        s.add_data_point(x, y)
    doc.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES, rdocx.Inches(5), rdocx.Inches(3), xy)
    doc.save(tmp_path / "c.docx")
    charts = rdocx.Document.open(tmp_path / "c.docx").charts
    assert [(c.chart_type, c.categories, [(x.name, x.values) for x in c.series]) for c in charts] == [
        (XL_CHART_TYPE.COLUMN_CLUSTERED, ("2024", "2025"), [("Inspections", (12.0, 15.0))]),
        (XL_CHART_TYPE.XY_SCATTER_LINES, ("0", "1"), [("Deflection", (1.5, 2.5))])]
    with zipfile.ZipFile(tmp_path / "c.docx") as z:
        assert any(n.startswith("word/embeddings/") and n.endswith(".xlsx") for n in z.namelist())   # editable
    assert run([rdocx_cli, "validate", tmp_path / "c.docx"]).returncode == 0
    before = doc.to_bytes()
    bad = CategoryChartData()
    bad.categories = ["a"]
    bad.add_series("one", (1,), number_format="0.0")
    bad.add_series("two", (2,), number_format="0%")
    with pytest.raises(ValueError):                                # one number format per chart
        doc.add_chart(XL_CHART_TYPE.PIE, rdocx.Inches(2), rdocx.Inches(2), bad)
    assert doc.to_bytes() == before
    assert not hasattr(charts[0], "replace_data")
    assert len(XL_CHART_TYPE) == 7


def test_embed_list_and_remove_a_font(tmp_path, free_font):
    doc = rdocx.Document()
    doc.add_paragraph("Embedded").runs[0].font.name = "Aileron"
    licence = rdocx.FontEmbeddingLicense(authorized=True, identity=FREE_FONT_LICENSE)
    doc.embed_font("Aileron", free_font, "regular", license=licence)
    doc.embed_font("Aileron", free_font.read_bytes(), "bold", license=licence)       # bytes, or a binary file
    assert doc.embed_true_type_fonts is True
    doc.save(tmp_path / "f.docx")
    reopened = rdocx.Document.open(tmp_path / "f.docx")
    font = next(f for f in reopened.fonts if f.name == "Aileron")
    assert [(e.kind, e.data == free_font.read_bytes(), e.license.authorized, e.license.identity, e.subsetted)
            for e in font.embedded_fonts] == [("regular", True, True, FREE_FONT_LICENSE, False),
                                              ("bold", True, True, FREE_FONT_LICENSE, False)]
    assert re.fullmatch(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}", font.embedded_fonts[0].font_key)
    assert "<w:embedTrueTypeFonts/>" in part(tmp_path / "f.docx", "word/settings.xml").decode()
    with zipfile.ZipFile(tmp_path / "f.docx") as z:
        stored = [z.read(n) for n in z.namelist() if n.endswith(".odttf")]
    assert len(stored) == 2 and free_font.read_bytes() not in stored                  # obfuscated as Word does
    removed = reopened.remove_embedded_font("Aileron", "bold")
    assert removed.kind == "bold" and [e.kind for e in next(f for f in reopened.fonts if f.name == "Aileron").embedded_fonts] == ["regular"]
    assert reopened.remove_embedded_font("Aileron", "italic") is None
    reopened.set_font("Aileron", family="swiss", pitch="variable")
    font = next(f for f in reopened.fonts if f.name == "Aileron")
    assert (font.family, font.pitch, len(font.embedded_fonts)) == ("swiss", "variable", 1)    # faces kept
    assert reopened.remove_font("Aileron").name == "Aileron" and reopened.remove_font("Aileron") is None
    reopened.save_subset_fonts = True
    reopened.embed_system_fonts = False
    assert (reopened.save_subset_fonts, reopened.embed_system_fonts) == (True, False)
    reopened.embed_true_type_fonts = None
    assert reopened.embed_true_type_fonts is None


def test_a_font_is_embedded_only_with_an_explicit_licence(tmp_path, free_font):
    doc = rdocx.Document()
    doc.add_paragraph("x")
    before = doc.to_bytes()
    for licence in (rdocx.FontEmbeddingLicense(authorized=False, identity="CC0-1.0: Aileron"),
                    rdocx.FontEmbeddingLicense(authorized=True, identity=" ")):
        with pytest.raises(rdocx.RdocxError):
            doc.embed_font("Aileron", free_font, license=licence)
    with pytest.raises(TypeError):
        doc.embed_font("Aileron", free_font)                                          # no licence, no embedding
    with pytest.raises(TypeError):
        rdocx.FontEmbeddingLicense(identity="CC0-1.0: Aileron")                       # authorized has no default
    for bad in (b"\x00\x01\x00\x00" + bytes(64), free_font.read_bytes().replace(b"name", b"nam_", 1)):
        with pytest.raises(rdocx.RdocxError):                                        # not a whole face
            doc.embed_font("Aileron", bad, license=rdocx.FontEmbeddingLicense(authorized=True, identity="CC0-1.0: Aileron"))
    ttc = b"ttcf" + free_font.read_bytes()[4:]
    with pytest.raises(rdocx.RdocxError):                                            # a collection is refused
        doc.embed_font("Aileron", ttc, license=rdocx.FontEmbeddingLicense(authorized=True, identity="CC0-1.0: Aileron"))
    assert digest(doc.to_bytes()) == digest(before)


def test_cli_font_commands(rdocx_cli, tmp_path, free_font):
    rdocx.Document().save(tmp_path / "d.docx")
    res = run([rdocx_cli, "font", "embed", tmp_path / "d.docx", free_font, "--name", "Aileron",
               "--license-identity", FREE_FONT_LICENSE, "-o", tmp_path / "no.docx"])
    assert res.returncode == 1 and "authorization" in res.stderr and not (tmp_path / "no.docx").exists()
    res = run([rdocx_cli, "font", "embed", tmp_path / "d.docx", free_font, "--name", "Aileron", "--face", "bold",
               "--license-authorized", "--license-identity", FREE_FONT_LICENSE, "-o", tmp_path / "e.docx", "--json"])
    assert res.returncode == 0, res.stderr
    record = json.loads(res.stdout)
    assert (record["name"], record["face"], record["byte_count"]) == ("Aileron", "bold", len(free_font.read_bytes()))
    listed = json.loads(run([rdocx_cli, "font", "list", "--json", tmp_path / "e.docx"], check=True).stdout)
    aileron = next(f for f in listed["fonts"] if f["name"] == "Aileron")
    assert [(e["face"], e["license_authorized"], e["license_identity"]) for e in aileron["embedded"]] == [
        ("bold", True, FREE_FONT_LICENSE)]
    assert listed["settings"]["embed_true_type_fonts"] is True
    assert "Aileron\tembedded: bold" in run([rdocx_cli, "font", "list", tmp_path / "e.docx"], check=True).stdout
    res = run([rdocx_cli, "font", "settings", tmp_path / "e.docx", "--save-subset-fonts", "true",
               "--embed-system-fonts", "false", "-o", tmp_path / "s.docx", "--json"])
    assert res.returncode == 0, res.stderr
    assert json.loads(res.stdout)["settings"] == {"embed_true_type_fonts": True, "embed_system_fonts": False,
                                                  "save_subset_fonts": True}
    assert run([rdocx_cli, "font", "remove", tmp_path / "s.docx", "--name", "Aileron", "--face", "bold",
                "-o", tmp_path / "r.docx"]).returncode == 0
    assert rdocx.Document.open(tmp_path / "r.docx").fonts[[f.name for f in rdocx.Document.open(tmp_path / "r.docx").fonts].index("Aileron")].embedded_fonts == ()
    assert run([rdocx_cli, "font", "remove", tmp_path / "s.docx", "--name", "Aileron", "-o", tmp_path / "g.docx"]).returncode == 0
    assert "Aileron" not in [f.name for f in rdocx.Document.open(tmp_path / "g.docx").fonts]
    res = run([rdocx_cli, "font", "remove", tmp_path / "s.docx", "--name", "Aileron", "-o", tmp_path / "r.docx"])
    assert res.returncode == 1 and "already exists" in res.stderr                                  # never overwrites


def test_what_the_new_edits_retire(free_font):
    """Font-table edits retire no handle; a chart or inserted HTML retires every handle, as a picture does."""
    from rdocx.chart.data import CategoryChartData
    from rdocx.enum.chart import XL_CHART_TYPE
    licence = rdocx.FontEmbeddingLicense(authorized=True, identity=FREE_FONT_LICENSE)
    data = CategoryChartData()
    data.categories = ["a"]
    data.add_series("s", (1,))
    edits = {
        "fonts": lambda d: (d.embed_font("Aileron", free_font, license=licence), d.set_font("Aileron", family="swiss"),
                            d.remove_embedded_font("Aileron", "regular"), d.remove_font("Aileron"),
                            setattr(d, "save_subset_fonts", True)),
        "chart": lambda d: d.add_chart(XL_CHART_TYPE.PIE, rdocx.Inches(1), rdocx.Inches(1), data),
        "html": lambda d: d.insert_html("<p>x</p>"),
    }
    retired = {}
    for name, edit in edits.items():
        doc = rdocx.Document()
        doc.add_paragraph("a").add_run("b")
        doc.add_table(1, 1)
        held = {"paragraph": doc.paragraphs[0], "run": doc.paragraphs[0].runs[0], "cell": doc.tables[0].cell(0, 0)}
        edit(doc)
        retired[name] = []
        for kind, handle in held.items():
            try:
                handle.text
            except rdocx.StaleElementError:
                retired[name].append(kind)
    assert retired == {"fonts": [], "chart": ["paragraph", "run", "cell"], "html": ["paragraph", "run", "cell"]}


def test_an_embedded_font_is_drawn_by_rdocx_and_libreoffice(rdocx_cli, tmp_path, free_font):
    """The font is installed nowhere: without the embedded face the render falls back to another font."""
    for embed in (False, True):
        doc = rdocx.Document()
        doc.add_paragraph("Embedded face text").runs[0].font.name = "Aileron"
        if embed:
            doc.embed_font("Aileron", free_font, license=rdocx.FontEmbeddingLicense(authorized=True, identity=FREE_FONT_LICENSE))
        doc.save(tmp_path / f"{embed}.docx")
        assert pdf_uses_font(rdocx.Document.open(tmp_path / f"{embed}.docx").to_pdf(), b"Aileron") is embed
    run([rdocx_cli, "convert", tmp_path / "True.docx", "--to", "pdf", "-o", tmp_path / "cli.pdf"], check=True)
    assert pdf_uses_font((tmp_path / "cli.pdf").read_bytes(), b"Aileron")
    assert pdf_uses_font(libreoffice_pdf(tmp_path / "True.docx", tmp_path), b"Aileron")


@pytest.mark.gap("html-space-inside-emphasis")
def test_the_space_before_emphasis_stays_outside_it():
    html = rdocx.Document.from_html("<p>see <u>this</u> and <s>that</s></p>")[0]
    md = rdocx.Document.from_markdown("a ~~struck~~ and *it*\n")[0]
    assert [(r.text, bool(r.underline), bool(r.font.strike)) for r in html.paragraphs[0].runs] == [
        ("see ", False, False), ("this", True, False), (" and ", False, False), ("that", False, True)]
    assert [(r.text, bool(r.font.strike), bool(r.italic)) for r in md.paragraphs[0].runs] == [
        ("a ", False, False), ("struck", True, False), (" and ", False, False), ("it", False, True)]


def test_a_restricted_fstype_is_not_checked(free_font):
    """The licence object is the only check: a face whose OS/2 fsType says Restricted License (0x0002) embeds."""
    data = bytearray(free_font.read_bytes())
    record = data.index(b"OS/2")
    offset = int.from_bytes(data[record + 8:record + 12], "big")
    data[offset + 8:offset + 10] = (2).to_bytes(2, "big")
    doc = rdocx.Document()
    doc.embed_font("Restricted", bytes(data), license=rdocx.FontEmbeddingLicense(authorized=True, identity="test"))
    prs = __import__("rpptx").Presentation()
    prs.embed_font("Restricted", bytes(data), license=__import__("rpptx").FontEmbeddingLicense(authorized=True, identity="test"))
    assert [f.name for f in doc.fonts if f.embedded_fonts] == ["Restricted"] and prs.embedded_fonts[0].typeface == "Restricted"


def test_the_space_before_emphasis_can_be_moved_out_after_the_import():
    """The workaround of gap html-space-inside-emphasis."""
    doc = rdocx.Document.from_html("<p>see <u>this</u> now</p>")[0]
    runs = doc.paragraphs[0].runs
    for k in range(1, len(runs)):
        if runs[k].text.startswith(" ") and (runs[k].underline or runs[k].font.strike):
            runs[k].text = runs[k].text[1:]
            runs[k - 1].text += " "
    assert [(r.text, bool(r.underline)) for r in doc.paragraphs[0].runs] == [("see ", False), ("this", True), (" now", False)]


DOUBLED_SPACES = {"<p>a <s>b </s> c</p>": "a b c", "<p>and <s> that</s></p>": "and that",
                  '<p>a <a href="https://example.com">b </a> c</p>': "a b c"}


@pytest.mark.gap("html-space-doubled-at-span-edge")
def test_spaces_on_both_sides_of_a_span_edge_collapse_into_one(rdocx_cli, tmp_path):
    """As a browser draws it: two collapsible spaces with a span edge between them are one space."""
    for html, text in DOUBLED_SPACES.items():
        assert rdocx.Document.from_html(html)[0].paragraphs[0].text == text, html
    (tmp_path / "page.html").write_text("<p>a <s>b </s> c</p>")
    run([rdocx_cli, "convert", tmp_path / "page.html", "-o", tmp_path / "page.docx"], check=True)
    assert rdocx.Document.open(tmp_path / "page.docx").paragraphs[0].text == "a b c"
    assert rdocx.Document.from_markdown("a [b ](https://example.com) c\n")[0].paragraphs[0].text == "a b c"
    doc = rdocx.Document()
    doc.insert_html("<p>a <s>b </s> c</p>")
    assert doc.paragraphs[-1].text == "a b c"


def test_a_doubled_space_can_be_dropped_after_the_import():
    """The workaround of gap html-space-doubled-at-span-edge."""
    for html, text in DOUBLED_SPACES.items():
        doc = rdocx.Document.from_html(html)[0]
        runs = doc.paragraphs[0].runs
        for k in range(1, len(runs)):
            if runs[k].text.startswith(" ") and runs[k - 1].text.endswith(" "):
                runs[k].text = runs[k].text[1:]
        assert doc.paragraphs[0].text == text, html
