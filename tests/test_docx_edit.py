"""docx, editing: counted replacement, structural edits, tables, pictures, handles, save semantics."""
import io
import os
import re
import zipfile

import docx
import pytest
import rdocx
from PIL import Image

from builders import every_story_docx, word_textbox_docx, wrapped_text_docx
from conftest import part, parts, run


def simple(path, *texts, table_after=None):
    d = docx.Document()
    for i, t in enumerate(texts):
        d.add_paragraph(t)
        if table_after == i:
            d.add_table(rows=2, cols=2).cell(0, 0).text = "cell"
    d.save(path)
    return path


# ---------------------------------------------------------------- replacement
def test_try_replace_text_counts_and_crosses_runs(tmp_path):
    d = docx.Document()
    p = d.add_paragraph()
    p.add_run("Hel")
    p.add_run("lo world")
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    assert doc.try_replace_text("Hello", "Bye") == 1
    assert doc.paragraphs[0].text == "Bye world"


def test_dry_run_through_bytes_leaves_document_untouched(report_docx):
    doc = rdocx.Document.open(report_docx)
    n = rdocx.Document.from_bytes(doc.to_bytes()).try_replace_text("footbridge", "FOOTBRIDGE")
    assert n > 1
    assert rdocx.Document.from_bytes(doc.to_bytes()).try_replace_text("footbridge", "FOOTBRIDGE") == n


def test_replace_all_regex(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha paragraph here.", "Beta paragraph after."))
    assert doc.replace_all_regex([(r"par\w+", "PARA")]) == 2
    assert [p.text for p in doc.paragraphs] == ["Alpha PARA here.", "Beta PARA after."]


def test_cli_replace_expect_refuses_and_writes_nothing(rdocx_cli, report_docx, tmp_path):
    out = tmp_path / "out.docx"
    bad = run([rdocx_cli, "replace", report_docx, "-p", "footbridge", "-v", "X", "--expect", "1", "-o", out])
    assert bad.returncode != 0 and not out.exists()
    good = run([rdocx_cli, "replace", report_docx, "-p", "harbour office", "-v", "port office", "--expect", "1", "-o", out])
    assert good.returncode == 0
    unchecked = run([rdocx_cli, "replace", report_docx, "-p", "footbridge", "-v", "X", "-o", tmp_path / "u.docx"])
    assert unchecked.returncode == 0  # without --expect, any count is written


def test_set_story_text_drops_the_paragraph_links(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("see ")
    doc.paragraphs[0].add_hyperlink("link", "https://example.org/")
    doc.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    doc.set_story_text(doc.story_items[0], "plain")
    doc.save(tmp_path / "b.docx")
    assert rdocx.Document.open(tmp_path / "b.docx").hyperlinks == ()
    assert b"w:hyperlink" not in part(tmp_path / "b.docx", "word/document.xml")


def test_replace_reaches_inline_and_block_content_controls(report_docx):
    doc = rdocx.Document.open(report_docx)
    assert doc.try_replace_text("74 out of 100", "75 out of 100") == 1
    assert doc.try_replace_text("harbour office", "port office") == 1


def test_python_replace_with_expected_count_is_all_or_nothing(report_docx):
    doc = rdocx.Document.open(report_docx)
    with pytest.raises(rdocx.RdocxError):
        doc.try_replace_text("footbridge", "X", expect=1)


def test_replacement_reach_by_story(tmp_path):
    """Replacement reaches the body, table cells, text boxes, headers and footers (once per variant part)."""
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    assert doc.try_replace_text("NEEDLE", "X") >= 6
    doc.save(tmp_path / "r.docx")
    out = parts(tmp_path / "r.docx")
    body = out["word/document.xml"]
    assert b"Body X one." in body and b"Cell X" in body and b"Text box X" in body
    notes = b"".join(b for n, b in out.items() if re.match(r"word/(header|footer)\d+\.xml$", n))
    assert b"Header X" in notes and b">Footer X" in notes and b"First footer X" in notes


@pytest.mark.gap("replace-footnotes")
def test_replace_reaches_footnotes(tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    doc.try_replace_text("NEEDLE", "X")
    doc.save(tmp_path / "r.docx")
    assert b"NEEDLE" not in part(tmp_path / "r.docx", "word/footnotes.xml")
    assert b"NEEDLE" not in part(tmp_path / "r.docx", "word/endnotes.xml")


@pytest.mark.gap("replace-tracked-insertions")
def test_replace_reaches_tracked_insertions(tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    assert "Tracked: ins NEEDLE" in [p.text for p in doc.paragraphs]
    doc.try_replace_text("ins NEEDLE", "ins X")
    assert "Tracked: ins X" in [p.text for p in doc.paragraphs]


def test_cli_replace_has_the_same_reach(rdocx_cli, tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    n = rdocx.Document.open(path).try_replace_text("NEEDLE", "X")
    assert run([rdocx_cli, "replace", path, "-p", "NEEDLE", "-v", "X", "--expect", n, "-o", tmp_path / "r.docx"]).returncode == 0


def test_replace_reaches_tables_in_headers_and_footers(tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    doc.try_replace_text("Footer cell NEEDLE", "Footer cell X")
    doc.save(tmp_path / "r.docx")
    footers = b"".join(b for n, b in parts(tmp_path / "r.docx").items() if n.startswith("word/footer"))
    assert b"Footer cell X" in footers


@pytest.mark.gap("textbox-alternate-content")
def test_word_text_box_is_read_and_counted_once(tmp_path):
    """Word writes a text box twice (DrawingML in mc:Choice, VML in mc:Fallback): one occurrence for a reader."""
    doc = rdocx.Document.open(word_textbox_docx(tmp_path / "w.docx", footnote=False))
    assert any(it.text == "Box NEEDLE" for it in doc.story_items)
    assert doc.try_replace_text("NEEDLE", "X") == 1


@pytest.mark.gap("text-wrapped-runs")
@pytest.mark.parametrize("wrapper", ["fldSimple", "smartTag", "customXml"])
def test_text_inside_simple_fields_smart_tags_and_custom_xml(rdocx_cli, tmp_path, wrapper):
    path = wrapped_text_docx(tmp_path / "w.docx", wrapper)
    assert rdocx.Document.open(path).paragraphs[0].text == "before MID after"


# ---------------------------------------------------------------- structure
def test_clone_insert_remove_pop(tmp_path):
    path = simple(tmp_path / "a.docx", "Alpha paragraph here.", "Beta paragraph after.")
    doc = rdocx.Document.open(path)
    doc.clone_content(doc.paragraphs[0], 1)
    assert [p.text for p in doc.paragraphs] == ["Alpha paragraph here."] * 2 + ["Beta paragraph after."]
    doc = rdocx.Document.open(path)
    doc.insert_paragraph(1, "Inserted")
    assert doc.paragraphs[1].text == "Inserted"
    doc.remove_content(1)
    frag = doc.pop_content(1)
    assert frag.kind == "paragraph"
    doc.insert_content(0, frag)
    assert [p.text for p in doc.paragraphs] == ["Beta paragraph after.", "Alpha paragraph here."]


def test_set_story_text_keeps_one_run(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha paragraph here."))
    doc.set_story_text(doc.story_items[0], "Replaced text")
    assert doc.paragraphs[0].text == "Replaced text"


KEEP = {
    "font": lambda d: setattr(d.paragraphs[1].runs[0].font, "bold", True),
    "paragraph_format": lambda d: setattr(d.paragraphs[1].paragraph_format, "space_after", rdocx.Pt(6)),
    "style": lambda d: setattr(d.paragraphs[1], "style", "Heading1"),
    "alignment": lambda d: setattr(d.paragraphs[1], "alignment", rdocx.WD_ALIGN_PARAGRAPH.CENTER),
    "numbering": lambda d: setattr(d.paragraphs[1], "numbering", (1, 0)),
    "cell_width": lambda d: setattr(d.tables[0].cell(0, 0), "width", rdocx.Inches(2)),
    "run_text": lambda d: setattr(d.paragraphs[1].runs[0], "text", "X"),
}
INVALIDATE = {
    "cell_text": lambda d: setattr(d.tables[0].cell(1, 1), "text", "Y"),
    "try_replace_text": lambda d: d.try_replace_text("delta", "D"),
    "replace_all_regex": lambda d: d.replace_all_regex([("delta", "D")]),
    "add_run": lambda d: d.paragraphs[1].add_run("more"),
    "clone_row": lambda d: d.tables[0].clone_row(0),
    "remove_row": lambda d: d.tables[0].remove_row(1),
    "add_comment": lambda d: d.add_comment(rdocx.RunRange(start=rdocx.RunPosition(body_index=0, run_index=0),
                                                          end=rdocx.RunPosition(body_index=0, run_index=1)), author="A", text="c"),
}


@pytest.mark.parametrize("op", sorted(KEEP) + sorted(INVALIDATE))
def test_which_calls_invalidate_handles(tmp_path, op):
    """The rule in python-api.md, "Handles": formatting setters and Run.text keep every handle valid, the
    other edits invalidate every handle (paragraph, run, table, story item)."""
    path = simple(tmp_path / "a.docx", "Alpha beta", "Gamma delta", table_after=1)
    doc = rdocx.Document.open(path)
    held = [lambda p=doc.paragraphs[0]: p.text, lambda r=doc.paragraphs[0].runs[0]: r.text,
            lambda t=doc.tables[0]: len(t.rows), lambda i=doc.story_items[0]: doc.set_story_text(i, "Alpha beta")]
    (KEEP.get(op) or INVALIDATE[op])(doc)
    for h in held:
        if op in KEEP:
            h()
        else:
            with pytest.raises(rdocx.StaleElementError):
                h()


def test_story_items_are_checked_handles(tmp_path):
    """A StoryItem taken before a structural change raises; re-fetch doc.story_items after each one."""
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "P0", "P1", "P2", "P3"))
    items = doc.story_items
    doc.set_story_text(items[1], "one")
    with pytest.raises(rdocx.StaleElementError):
        doc.set_story_text(items[2], "two")
    doc.set_story_text(doc.story_items[2], "two")
    assert [p.text for p in doc.paragraphs] == ["P0", "one", "two", "P3"]


def test_move_content_destination_counts_before_the_move(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", *"ABCDE"))
    doc.move_content(doc.paragraphs[1], 3)
    assert "".join(p.text for p in doc.paragraphs) == "ACBDE"


def test_clone_content_renames_bookmarks_and_drops_comment_anchors(tmp_path):
    from docx.oxml.ns import qn
    d = docx.Document()
    p = d.add_paragraph()
    p.add_run("Cloned text")
    p._p.insert(1, p._p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): "7", qn("w:name"): "mark"}))
    p._p.append(p._p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): "7"}))
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    rng = rdocx.RunRange(start=rdocx.RunPosition(body_index=0, run_index=0), end=rdocx.RunPosition(body_index=0, run_index=1))
    doc.add_comment(rng, author="A", text="on the source")
    doc.clone_content(doc.paragraphs[0], 1)
    doc.save(tmp_path / "b.docx")
    xml = part(tmp_path / "b.docx", "word/document.xml").decode()
    names = re.findall(r'<w:bookmarkStart [^>]*w:name="([^"]+)"', xml)
    assert xml.count("<w:commentRangeStart") == 1
    assert names[0] == "mark" and len(names) == 2 and names[1] != "mark"


def test_handles_go_stale_after_structural_edits(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha", "Beta"))
    held = doc.paragraphs[0]
    doc.insert_paragraph(0, "x")
    with pytest.raises(rdocx.StaleElementError):
        held.text


def test_split_run_then_format_a_word(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha", "Beta paragraph after."))
    doc.split_run(1, 0, 4)
    doc.paragraphs[1].runs[0].font.bold = True
    assert [r.text for r in doc.paragraphs[1].runs] == ["Beta", " paragraph after."]
    assert doc.paragraphs[1].runs[0].font.bold is True


def test_split_run_body_index_after_a_table(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha", "Beta paragraph after.", "Gamma.", table_after=0))
    bi = doc.find_content_index(doc.paragraphs[1])
    doc.split_run(bi, 0, 4)
    assert [r.text for r in doc.paragraphs[1].runs] == ["Beta", " paragraph after."]


@pytest.mark.gap("docx-paragraph-text-setter")
def test_paragraph_text_setter(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha"))
    doc.paragraphs[0].text = "Omega"
    assert doc.paragraphs[0].text == "Omega"


# ---------------------------------------------------------------- tables
def test_table_clone_rewrite_remove_row(report_docx):
    doc = rdocx.Document.open(report_docx)
    n = len(doc.tables[0].rows)
    doc.tables[0].clone_row(n - 1)
    for c in range(len(doc.tables[0].rows[n].cells)):
        doc.tables[0].rows[n].cells[c].text = "cloned"
    assert len(doc.tables[0].rows) == n + 1 and doc.tables[0].cell(n, 1).text == "cloned"
    doc.tables[0].remove_row(n)
    assert len(doc.tables[0].rows) == n


def test_row_identity_attributes_survive_an_edit(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    before = len(re.findall(rb"<w:tr [^>]*w14:paraId", part(report_docx, "word/document.xml")))
    after = len(re.findall(rb"<w:tr [^>]*w14:paraId", part(tmp_path / "e.docx", "word/document.xml")))
    assert before == after == 71


def test_cell_and_table_widths_are_settable(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x", table_after=0))
    doc.tables[0].cell(0, 0).width = rdocx.Inches(4)
    doc.tables[0].width = rdocx.Inches(5)
    doc.save(tmp_path / "b.docx")
    t = docx.Document(tmp_path / "b.docx").tables[0]
    assert abs(t.cell(0, 0).width.inches - 4) < 0.01


def test_table_merge_borders_shading_widths_and_row_height_from_python(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x", table_after=0))
    doc.tables[0].cell(1, 1).text = "full"
    with pytest.raises(rdocx.RdocxError):
        doc.tables[0].set_cell_grid_span(1, 0, 2)  # a span consumes empty cells only
    t = doc.tables[0]
    t.set_cell_grid_span(0, 0, 2)  # consumes the empty cell to its right; a structural change
    with pytest.raises(rdocx.StaleElementError):
        t.cell(0, 0)
    t = doc.tables[0]
    t.cell(1, 1).shading = "FFEE00"
    t.set_borders("single", size=4, color="000000")
    t.set_column_width(0, rdocx.Inches(1))
    t.rows[1].height = rdocx.Inches(0.5)
    doc.save(tmp_path / "b.docx")
    x = part(tmp_path / "b.docx", "word/document.xml").decode()
    assert [len(r.cells) for r in docx.Document(tmp_path / "b.docx").tables[0].rows] == [2, 2]
    assert '<w:gridSpan w:val="2"/>' in x and 'w:fill="FFEE00"' in x and "<w:tblBorders>" in x
    assert '<w:gridCol w:w="1440"/>' in x and '<w:trHeight w:val="720"' in x


# ---------------------------------------------------------------- pictures, hyperlinks, sections, styles
def first_embed(doc):
    for item in doc.story_items:
        xml = item.xml.decode() if isinstance(item.xml, bytes) else (item.xml or "")
        m = re.search(r'r:embed="([^"]+)"', xml)
        if m:
            return m.group(1)


def test_replace_image_keeps_extent(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    rid = first_embed(doc)
    size = Image.open(io.BytesIO(doc.image_data(rid))).size
    png = io.BytesIO()
    Image.new("RGB", size, (90, 120, 160)).save(png, "PNG")
    before = re.findall(rb'<wp:extent [^>]*>', part(report_docx, "word/document.xml"))
    doc.replace_image(rid, png.getvalue())
    doc.save(tmp_path / "p.docx")
    assert re.findall(rb'<wp:extent [^>]*>', part(tmp_path / "p.docx", "word/document.xml")) == before
    assert rdocx.Document.open(tmp_path / "p.docx").image_data(rid) == png.getvalue()


def with_content_control(path, default_ns):
    """A paragraph, then a block content control (Google Docs style), optionally a default namespace on the
    root (as Google Docs exports do)."""
    from docx.oxml.ns import qn
    d = docx.Document()
    d.add_paragraph("Before the control.")
    body = d.element.body
    sdt = body.makeelement(qn("w:sdt"), {})
    pr = sdt.makeelement(qn("w:sdtPr"), {})
    tag = pr.makeelement(qn("w:tag"), {})
    tag.set(qn("w:val"), "goog_rdk_0")
    pr.append(tag)
    sdt.append(pr)
    content = sdt.makeelement(qn("w:sdtContent"), {})
    p = d.add_paragraph("Inside the control.")._p
    body.remove(p)
    content.append(p)
    sdt.append(content)
    body.insert(1, sdt)
    d.save(path)
    if default_ns:
        items = [(i, zipfile.ZipFile(path).read(i.filename)) for i in zipfile.ZipFile(path).infolist()]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for info, data in items:
                if info.filename == "word/document.xml":
                    data = data.replace(b"<w:document ", b'<w:document xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" ', 1)
                z.writestr(info, data)
    return path


def png(size=(60, 40)):
    buf = io.BytesIO()
    Image.new("RGB", size).save(buf, "PNG")
    return buf.getvalue()


def test_add_picture_with_a_content_control(tmp_path):
    doc = rdocx.Document.open(with_content_control(tmp_path / "a.docx", default_ns=False))
    doc.add_picture(png(), "x.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
    doc.save(tmp_path / "b.docx")
    assert len(docx.Document(tmp_path / "b.docx").inline_shapes) == 1


def test_add_picture_with_a_content_control_and_a_default_namespace(tmp_path):
    doc = rdocx.Document.open(with_content_control(tmp_path / "a.docx", default_ns=True))
    doc.add_picture(png(), "x.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
    doc.save(tmp_path / "b.docx")


@pytest.mark.gap("docx-picture-resize")
def test_resize_existing_picture(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.set_picture_size("rId1", width=rdocx.Inches(1), height=rdocx.Inches(1))


@pytest.mark.gap("docx-hyperlink-retarget")
def test_retarget_hyperlink(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.set_hyperlink_url(doc.hyperlinks[0], "https://example.org/new")


def test_section_margins_are_writable(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.update_section(0, margin_top=rdocx.Inches(0.5))  # Section itself is a read-only snapshot
    assert rdocx.Document.from_bytes(doc.to_bytes()).sections[0].margin_top == rdocx.Inches(0.5)


@pytest.mark.gap("docx-python-styles")
def test_create_paragraph_style(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.add_style("Note", style_type="paragraph", based_on="Normal")


def test_bookmarks_from_python(report_docx):
    assert rdocx.Document.open(report_docx).bookmarks is not None


@pytest.mark.gap("style-id-unchecked")
def test_unknown_style_id_is_refused(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x"))
    with pytest.raises((rdocx.RdocxError, ValueError, KeyError)):
        doc.paragraphs[0].style = "NoSuchStyle"


@pytest.mark.gap("docx-core-properties")
def test_core_properties_are_writable(tmp_path):
    doc = rdocx.Document()
    doc.core_properties.title = "A title"
    doc.save(tmp_path / "a.docx")
    assert docx.Document(tmp_path / "a.docx").core_properties.title == "A title"


# ---------------------------------------------------------------- formatting values and units
def test_font_color_reads_back_an_rgbcolor(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("x")
    doc.paragraphs[0].runs[0].font.color = rdocx.RGBColor(0x7B, 0x1E, 0x3A)
    color = doc.paragraphs[0].runs[0].font.color
    assert tuple(color) == (0x7B, 0x1E, 0x3A) and str(color) == "7B1E3A" and color != "7B1E3A"
    with pytest.raises(TypeError):
        doc.paragraphs[0].runs[0].font.color = "7B1E3A"


def test_length_helpers():
    assert rdocx.Pt(1) == 12700 and rdocx.Inches(1) == 914400 and rdocx.Cm(1) == 360000
    assert rdocx.Mm(1) == 36000 and rdocx.Emu(5) == 5 and rdocx.Inches(1).pt == 72


# ---------------------------------------------------------------- save semantics
@pytest.mark.gap("template-save-as-document")
def test_template_saved_as_document_gets_the_document_content_type(tmp_path):
    src = simple(tmp_path / "a.docx", "x")
    with zipfile.ZipFile(src) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(tmp_path / "t.dotx", "w") as z:
        for info, data in items:
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"document.main+xml", b"template.main+xml")
            z.writestr(info, data)
    rdocx.Document.open(tmp_path / "t.dotx").save(tmp_path / "b.docx")
    assert b"wordprocessingml.document.main+xml" in part(tmp_path / "b.docx", "[Content_Types].xml")


def test_noop_save_is_byte_identical(report_docx, tmp_path):
    rdocx.Document.open(report_docx).save(tmp_path / "n.docx")
    a, b = parts(report_docx), parts(tmp_path / "n.docx")
    assert [n for n in a if a[n] != b.get(n)] == []


def test_save_keeps_ignorable_prefixes_declared(report_docx, tmp_path):
    rdocx.Document.open(report_docx).save(tmp_path / "n.docx")
    root = part(tmp_path / "n.docx", "word/comments.xml").decode().split(">", 2)[1]
    ignorable = re.search(r'mc:Ignorable="([^"]*)"', root).group(1).split()
    assert all(f"xmlns:{p}=" in root for p in ignorable)


@pytest.mark.parametrize("story", ["body", "footer"])
def test_edit_keeps_the_root_ignorable(tmp_path, story):
    """python-docx's document and footer roots declare w14 and wp14 and list them in mc:Ignorable. An edit that
    rewrites the part keeps the attribute, the declarations and any w14:paraId in the part."""
    d = docx.Document()
    d.add_paragraph("Body alpha.")
    d.sections[0].footer.paragraphs[0].text = "Footer beta."
    d.save(tmp_path / "a.docx")
    name = {"body": "word/document.xml", "footer": "word/footer1.xml"}[story]
    assert b'mc:Ignorable="w14 wp14"' in part(tmp_path / "a.docx", name)
    doc = rdocx.Document.open(tmp_path / "a.docx")
    assert doc.try_replace_text({"body": "alpha", "footer": "beta"}[story], "gamma") == 1
    doc.save(tmp_path / "b.docx")
    root = part(tmp_path / "b.docx", name).decode().split(">", 2)[1]
    assert "xmlns:w14=" in root and 'mc:Ignorable="w14 wp14"' in root


def test_save_replaces_the_file_atomically(tmp_path):
    path = simple(tmp_path / "a.docx", "Alpha")
    before = os.stat(path).st_ino
    rdocx.Document.open(path).save(path)
    assert os.stat(path).st_ino != before


def test_edit_keeps_package_readable_by_python_docx(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    x = docx.Document(tmp_path / "e.docx")
    assert any("outlined" in p.text for p in x.paragraphs)
    with zipfile.ZipFile(tmp_path / "e.docx") as z:
        assert z.namelist()[0] == "[Content_Types].xml"


# ---------------------------------------------------------------- other stories and settings
def test_header_footer_hyperlink_and_update_fields_on_open(tmp_path):
    doc = rdocx.Document()
    doc.add_paragraph("x")
    doc.set_header("Head")
    doc.set_footer("Foot")
    footer = next(s for s in doc.stories if s.kind == "footer")
    doc.add_hyperlink_to_story(footer, "site", "https://example.org/f")
    doc.update_fields_on_open = True
    doc.save(tmp_path / "a.docx")
    again = rdocx.Document.open(tmp_path / "a.docx")
    assert [h.url for h in again.hyperlinks] == ["https://example.org/f"]
    assert b"updateFields" in part(tmp_path / "a.docx", "word/settings.xml")


def test_set_footer_replaces_the_fields_too(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    assert b"PAGE" in part(report_docx, "word/footer1.xml")
    doc.set_footer("Plain footer")
    doc.save(tmp_path / "f.docx")
    footer = part(tmp_path / "f.docx", "word/footer1.xml")
    assert b"Plain footer" in footer and b"fldChar" not in footer and b"fldSimple" not in footer
