"""docx, editing: counted replacement, structural edits, tables, pictures, handles, save semantics."""
import io
import os
import re
import zipfile

import docx
import pytest
import rdocx
from PIL import Image

from builders import every_story_docx, word_textbox_docx, wrapped_run_docx, wrapped_text_docx
from conftest import digest, part, parts, run


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


def test_replace_reaches_footnotes(tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    doc.try_replace_text("NEEDLE", "X")
    doc.save(tmp_path / "r.docx")
    assert b"NEEDLE" not in part(tmp_path / "r.docx", "word/footnotes.xml")
    assert b"NEEDLE" not in part(tmp_path / "r.docx", "word/endnotes.xml")


def test_replace_reaches_tracked_insertions(tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    assert "Tracked: ins NEEDLE" in [p.text for p in doc.paragraphs]
    doc.try_replace_text("ins NEEDLE", "ins X")
    assert "Tracked: ins X" in [p.text for p in doc.paragraphs]
    assert [r.kind for r in doc.revisions].count("insertion") == 1  # the new text stays a tracked insertion
    doc.reject_all()
    assert "Tracked: del NEEDLE" in [p.text for p in doc.paragraphs]  # the insertion goes, the deletion comes back


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


def test_word_text_box_is_read_and_counted_once(tmp_path):
    """Word writes a text box twice (DrawingML in mc:Choice, VML in mc:Fallback): one occurrence for a reader."""
    doc = rdocx.Document.open(word_textbox_docx(tmp_path / "w.docx", footnote=False))
    assert any(it.text == "Box NEEDLE" for it in doc.story_items)
    assert doc.try_replace_text("NEEDLE", "X") == 1


@pytest.mark.parametrize("wrapper", ["fldSimple", "smartTag", "customXml"])
def test_text_inside_simple_fields_smart_tags_and_custom_xml(rdocx_cli, tmp_path, wrapper):
    path = wrapped_text_docx(tmp_path / "w.docx", wrapper)
    assert rdocx.Document.open(path).paragraphs[0].text == "before MID after"


@pytest.mark.parametrize("wrapper", ["sdt", "ins", "fldSimple", "smartTag"])
def test_a_match_across_a_wrapper_edge_is_not_replaced(tmp_path, wrapper):
    """A match inside a content control, an insertion, a simple field or a smart tag is replaced there; one
    that crosses its edge is not (docx_ops.replace_batch then refuses: the text is still there)."""
    if wrapper in ("sdt", "ins"):
        path = wrapped_run_docx(tmp_path / "w.docx", wrapper, target="MID")
    else:
        path = wrapped_text_docx(tmp_path / "w.docx", wrapper)
    doc = rdocx.Document.open(path)
    assert doc.paragraphs[0].text == "before MID after"
    assert doc.try_replace_text("before MID", "X") == 0
    assert doc.try_replace_text("MID", "X") == 1 and doc.paragraphs[0].text == "before X after"


def test_paragraph_replace_text_is_scoped_to_its_paragraph(tmp_path):
    """The same clause in two paragraphs: only the target one changes, across runs, its comment kept; a wrong
    `expect` raises ReplacementCountError and changes nothing."""
    d = docx.Document()
    p = d.add_paragraph()
    p.add_run("the old ")
    p.add_run("clause").bold = True
    p.add_run(" here")
    d.add_paragraph("the old clause there")
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    doc.add_comment(rdocx.RunRange(start=rdocx.RunPosition(body_index=0, run_index=0),
                                   end=rdocx.RunPosition(body_index=0, run_index=3)), author="R", text="c")
    assert rdocx.Document.from_bytes(doc.to_bytes()).try_replace_text("old clause", "x") == 2  # the document-wide call
    assert doc.paragraphs[0].replace_text("old clause", "new clause", expect=1) == 1
    assert [p.text for p in doc.paragraphs] == ["the new clause here", "the old clause there"]
    assert doc.comments[0].anchor_text == "the new clause here"
    with pytest.raises(rdocx.ReplacementCountError):
        doc.paragraphs[1].replace_text("old clause", "x", expect=2)
    assert issubclass(rdocx.ReplacementCountError, rdocx.RdocxError)
    assert [p.text for p in doc.paragraphs] == ["the new clause here", "the old clause there"]
    assert doc.paragraphs[1].replace_text("absent", "x") == 0


def test_cell_replace_text_is_scoped_to_its_cell(tmp_path):
    d = docx.Document()
    d.add_paragraph("No action in the body")
    t = d.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = "No action"
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    assert doc.tables[0].cell(1, 1).replace_text("No action", "Monitor", expect=1) == 1
    assert [c.text for row in doc.tables[0].rows for c in row.cells] == ["No action"] * 3 + ["Monitor"]
    with pytest.raises(rdocx.ReplacementCountError):
        doc.tables[0].cell(0, 0).replace_text("No action", "Monitor", expect=2)
    assert doc.tables[0].cell(0, 0).text == "No action" and doc.paragraphs[0].text == "No action in the body"
    cell_paragraph = doc.tables[0].cell(0, 1).paragraphs[0]                 # a paragraph of a cell works too
    assert cell_paragraph.replace_text("action", "change", expect=1) == 1 and doc.tables[0].cell(0, 1).text == "No change"


def test_replace_story_text_edits_one_item_of_another_story(tmp_path):
    d = docx.Document()
    d.add_paragraph("NEEDLE in the body")
    d.add_table(rows=1, cols=2).cell(0, 1).text = "NEEDLE in a cell, NEEDLE"
    d.sections[0].header.paragraphs[0].text = "NEEDLE in the header"
    d.sections[0].footer.paragraphs[0].text = "NEEDLE in the footer"
    d.save(tmp_path / "h.docx")
    doc = rdocx.Document.open(tmp_path / "h.docx")

    def item(story, kind="paragraph"):                                     # the last one: cell (0, 1) for a cell
        return [i for i in doc.story_items if i.story.kind == story and i.kind == kind][-1]
    for story in ("header", "footer"):
        assert doc.replace_story_text(item(story), "NEEDLE", "PIN", expect=1) == 1
    assert doc.replace_story_text(item("table_cell"), "NEEDLE", "PIN", expect=2) == 2
    assert doc.paragraphs[0].text == "NEEDLE in the body"
    assert [item(s).text for s in ("header", "footer", "table_cell")] == [
        "PIN in the header", "PIN in the footer", "PIN in a cell, PIN"]
    doc.tables[0].cell(0, 0).text = "NEEDLE"
    assert doc.replace_story_text(item("body", "table"), "NEEDLE", "PIN", expect=1) == 1   # a whole body table
    assert doc.tables[0].cell(0, 0).text == "PIN" and doc.paragraphs[0].text == "NEEDLE in the body"
    with pytest.raises(rdocx.ReplacementCountError):
        doc.replace_story_text(doc.story_items[0], "NEEDLE", "PIN", expect=3)
    doc.add_comment_on_text("NEEDLE", author="R", text="NEEDLE here too")
    comment = next(i for i in doc.story_items if i.story.kind == "comment" and i.kind == "paragraph")
    with pytest.raises(rdocx.RdocxError):                                 # refused: comments are never searched
        doc.replace_story_text(comment, "NEEDLE", "PIN")
    box =rdocx.Document.open(word_textbox_docx(tmp_path / "t.docx"))
    note = next(i for i in box.story_items if i.story.kind == "footnote" and i.kind == "paragraph")
    assert box.replace_story_text(note, "NEEDLE", "PIN", expect=1) == 1
    text_box = next(i for i in box.story_items if i.story.kind == "text_box" and i.kind == "paragraph")
    with pytest.raises(rdocx.RdocxError):                                 # refused: the document-wide call edits both copies
        box.replace_story_text(text_box, "NEEDLE", "PIN")
    assert box.try_replace_text("NEEDLE", "PIN") == 1                     # the text box, counted once


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
    "paragraph_text": lambda d: setattr(d.paragraphs[1], "text", "Z"),
    "try_replace_text": lambda d: d.try_replace_text("delta", "D"),
    "replace_all_regex": lambda d: d.replace_all_regex([("delta", "D")]),
    "add_run": lambda d: d.paragraphs[1].add_run("more"),
    "clone_row": lambda d: d.tables[0].clone_row(0),
    "remove_row": lambda d: d.tables[0].remove_row(1),
    "paragraph_replace_text": lambda d: d.paragraphs[1].replace_text("delta", "D"),
    "cell_replace_text": lambda d: d.tables[0].cell(0, 0).replace_text("cell", "C"),
    "replace_story_text": lambda d: d.replace_story_text(d.story_items[1], "delta", "D"),
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


def test_paragraph_text_setter(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "Alpha"))
    doc.paragraphs[0].text = "Omega"
    assert doc.paragraphs[0].text == "Omega"


def test_paragraph_text_setter_keeps_format_and_comments(tmp_path):
    """One run without direct formatting; the paragraph style and format, and comment anchors, stay."""
    d = docx.Document()
    p = d.add_paragraph(style="Heading 1")
    p.add_run("Bold").bold = True
    p.add_run(" plain")
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    doc.paragraphs[0].paragraph_format.keep_together = True
    doc.add_comment(rdocx.RunRange(start=rdocx.RunPosition(body_index=0, run_index=0),
                                   end=rdocx.RunPosition(body_index=0, run_index=2)), author="R", text="c")
    doc.paragraphs[0].text = "New title"
    doc.save(tmp_path / "b.docx")
    p = docx.Document(tmp_path / "b.docx").paragraphs[0]
    assert (p.text, p.style.name, p.paragraph_format.keep_together) == ("New title", "Heading 1", True)
    assert (p.runs[0].text, p.runs[0].bold) == ("New title", None)
    assert b"commentRangeStart" in part(tmp_path / "b.docx", "word/document.xml")


def test_paragraph_text_setter_refuses_part_of_a_toc(report_docx):
    doc = rdocx.Document.open(report_docx)
    i = next(k for k, p in enumerate(doc.paragraphs) if p.style == "TOC1")  # holds the TOC field's begin
    with pytest.raises((rdocx.RdocxError, ValueError)):
        doc.paragraphs[i].text = "x"


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


WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
DOCPR_ID = re.compile(rb'(<wp:docPr\b[^>]*?\sid=")(\d+)(")')


def two_pictures(ids):
    """Two inline pictures in the body, with these two wp:docPr ids."""
    d = docx.Document()
    for name in "AB":
        d.add_paragraph(f"Picture {name}: ").add_run().add_picture(io.BytesIO(png()), width=docx.shared.Inches(1))
    for docpr, value in zip(d.element.body.iter(WP + "docPr"), ids):
        docpr.set("id", str(value))
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def docpr_ids(xml):
    return [int(m[2]) for m in DOCPR_ID.finditer(xml)]


@pytest.mark.parametrize("ids", [(1, 2), (0, 1), (0, 0), (7, 7)])
def test_drawings_of_one_part_that_share_an_id(ids, rdocx_cli, tmp_path):
    """Two pictures of the body with one wp:docPr id (one id in two parts already opens): the file opens in the CLI
    and the binding, a new picture gets an id that no other drawing of the part has, and the saved file opens."""
    f = tmp_path / "a.docx"
    f.write_bytes(two_pictures(ids))
    assert run([rdocx_cli, "text", f]).returncode == 0
    doc = rdocx.Document.open(f)
    doc.add_picture(png(), "x.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
    doc.save(tmp_path / "b.docx")
    got = docpr_ids(part(tmp_path / "b.docx", "word/document.xml"))
    assert len(got) == 3 and got.count(got[-1]) == 1
    assert len(docx.Document(tmp_path / "b.docx").inline_shapes) == 3
    rdocx.Document.open(tmp_path / "b.docx")


def test_resize_existing_picture(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    drawing = next(it for it in doc.story_items if it.kind == "drawing")
    xml = drawing.xml.decode() if isinstance(drawing.xml, bytes) else drawing.xml
    rid = re.search(r'r:embed="([^"]+)"', xml).group(1)
    assert doc.set_picture_size(rid, width=rdocx.Inches(1), height=rdocx.Inches(2)) >= 1
    doc.save(tmp_path / "b.docx")
    sizes = [(s.width, s.height) for s in docx.Document(tmp_path / "b.docx").inline_shapes]
    assert (rdocx.Inches(1), rdocx.Inches(2)) in sizes


def test_retarget_and_remove_hyperlinks(report_docx):
    """Retargeting keeps every Hyperlink snapshot valid; a removal makes the older ones of its story raise."""
    doc = rdocx.Document.open(report_docx)
    links = doc.hyperlinks
    text = links[0].text
    for k, link in enumerate(links):
        doc.set_hyperlink_url(link, f"https://example.org/{k}")
    assert [h.url for h in doc.hyperlinks] == [f"https://example.org/{k}" for k in range(len(links))]
    stale = doc.hyperlinks
    doc.remove_hyperlink(stale[0])
    assert len(doc.hyperlinks) == len(stale) - 1 and any(text in p.text for p in doc.paragraphs)
    same = next(h for h in stale[1:] if (h.story.kind, h.story.part_name) == (stale[0].story.kind, stale[0].story.part_name))
    with pytest.raises(rdocx.RdocxError):
        doc.set_hyperlink_url(same, "https://example.org/x")


def test_section_margins_are_writable(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.update_section(0, margin_top=rdocx.Inches(0.5))  # Section itself is a read-only snapshot
    assert rdocx.Document.from_bytes(doc.to_bytes()).sections[0].margin_top == rdocx.Inches(0.5)


def test_create_paragraph_style(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.add_style("Note", style_type="paragraph", based_on="Normal")
    doc.paragraphs[0].style = "Note"
    assert doc.paragraphs[0].style == "Note"


def test_set_style_changes_only_what_it_is_given(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x"))
    doc.add_style("Note box", based_on="Normal", italic=True, left_indent=rdocx.Inches(0.5))
    doc.set_style("Notebox", bold=True, space_after=rdocx.Pt(6))  # by id or by name
    doc.set_style("Heading 1", font_name="Arial", color=rdocx.RGBColor(0x11, 0x22, 0x33))
    back = docx.Document(io.BytesIO(doc.to_bytes()))
    note = back.styles["Note box"]
    assert (note.font.bold, note.font.italic) == (True, True)
    assert note.paragraph_format.left_indent == rdocx.Inches(0.5)
    assert note.paragraph_format.space_after == rdocx.Pt(6)
    assert note.base_style.style_id == "Normal"
    # a theme font and a theme colour the style already has stay, and Word uses them over the new values
    rpr = part(io.BytesIO(doc.to_bytes()), "word/styles.xml").decode()
    heading = re.search(r'<w:style [^>]*w:styleId="Heading1".*?</w:style>', rpr, re.S).group(0)
    assert 'w:ascii="Arial"' in heading and "w:asciiTheme=" in heading
    assert 'w:val="112233"' in heading and "w:themeColor=" in heading
    before = doc.to_bytes()
    with pytest.raises(KeyError):
        doc.set_style("No such style", bold=True)
    assert digest(doc.to_bytes()) == digest(before)


def test_bookmarks_from_python(report_docx):
    assert rdocx.Document.open(report_docx).bookmarks is not None


def test_unknown_style_id_is_refused(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x"))
    with pytest.raises((rdocx.RdocxError, ValueError, KeyError)):
        doc.paragraphs[0].style = "NoSuchStyle"


def test_style_is_assigned_by_id_or_name(tmp_path):
    doc = rdocx.Document.open(simple(tmp_path / "a.docx", "x"))
    doc.paragraphs[0].style = "Heading 1"
    assert doc.paragraphs[0].style == "Heading1"
    with pytest.raises(KeyError):
        doc.paragraphs[0].style = "NoSuchStyle"


def google_styles_docx(duplicate_ids=True, several_defaults=False):
    """A TOC, a heading, and a styles part with what Google Docs writes: with `duplicate_ids`, repeated style
    ids (`TableNormal` three times, a second `Normal` with another body, `Table1` twice with different bodies);
    with `several_defaults`, a second default table style (`TableauNormal`) and a second default paragraph
    style (`NormalWeb`) under ids of their own."""
    import copy

    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    d = docx.Document()
    p = d.add_paragraph()
    for xml in ('<w:fldChar w:fldCharType="begin"/>', '<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText>',
                '<w:fldChar w:fldCharType="separate"/>', '<w:t>Old entry</w:t>'):
        p._p.append(parse_xml(f"<w:r {nsdecls('w')}>{xml}</w:r>"))
    d.add_paragraph()._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))
    d.add_paragraph("Scope", style="Heading 1")
    styles = d.styles.element
    if duplicate_ids:
        first = next(s for s in styles.findall(qn("w:style")) if s.get(qn("w:styleId")) == "TableNormal")
        styles.append(copy.deepcopy(first))
        styles.append(copy.deepcopy(first))
        styles.append(parse_xml(f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="Normal"><w:name w:val="normal"/>'
                                '<w:rPr><w:sz w:val="30"/></w:rPr></w:style>'))
        for body in ("", '<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>'):
            styles.append(parse_xml(f'<w:style {nsdecls("w")} w:type="table" w:styleId="Table1"><w:name w:val="Table1"/>{body}</w:style>'))
    if several_defaults:
        for kind, sid, name in (("table", "TableauNormal", "Tableau Normal"), ("paragraph", "NormalWeb", "Normal (Web)")):
            styles.append(parse_xml(f'<w:style {nsdecls("w")} w:type="{kind}" w:default="1" w:styleId="{sid}">'
                                    f'<w:name w:val="{name}"/></w:style>'))
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def default_style_ids(data):
    from lxml import etree
    from docx.oxml.ns import qn
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/styles.xml"))
    return sorted((s.get(qn("w:type")), s.get(qn("w:styleId"))) for s in root.findall(qn("w:style"))
                  if s.get(qn("w:default")) in ("1", "true", "on"))


def test_add_style_with_duplicate_style_ids():
    doc = rdocx.Document.from_bytes(google_styles_docx())
    doc.add_style("Note box", based_on="Normal")  # the first Normal is authoritative, as for rebuild_toc
    back = docx.Document(io.BytesIO(doc.to_bytes()))
    assert back.styles["Note box"].base_style.style_id == "Normal"


def test_add_style_with_several_default_styles_of_one_type():
    doc = rdocx.Document.from_bytes(google_styles_docx(duplicate_ids=False, several_defaults=True))
    doc.add_style("Note box", based_on="Normal")  # the first default of each type is authoritative
    back = docx.Document(io.BytesIO(doc.to_bytes()))
    assert back.styles["Note box"].base_style.style_id == "Normal"


def test_rebuild_toc_uses_the_first_of_duplicate_style_ids():
    doc = rdocx.Document.from_bytes(google_styles_docx())
    report = doc.rebuild_toc()
    assert report.entry_count == 1
    assert "duplicate style ID 'TableNormal' used first definition while rebuilding TOC" in report.diagnostics


def test_rebuild_toc_accepts_several_default_styles_of_one_type():
    doc = rdocx.Document.from_bytes(google_styles_docx(duplicate_ids=False, several_defaults=True))
    assert doc.rebuild_toc().entry_count == 1


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


def test_edit_rewrites_only_the_part_it_touches(report_docx, tmp_path):
    """A body replacement rewrites word/document.xml; styles, numbering, headers, footers, comments and media keep
    their bytes, and the package gains no part."""
    doc = rdocx.Document.open(report_docx)
    assert doc.try_replace_text("inspection", "survey") > 0
    doc.save(tmp_path / "e.docx")
    a, b = parts(report_docx), parts(tmp_path / "e.docx")
    assert sorted(b) == sorted(a)
    assert [n for n in a if a[n] != b[n]] == ["word/document.xml"]


def test_edit_rewrites_its_part_in_the_source_layout(tmp_path):
    """Word writes w:rsid* on nearly every paragraph and run, and no indentation. An edit of one paragraph keeps
    the rest of document.xml as it was: xmlns:w declared once, no line added, the size within a few bytes."""
    d = docx.Document()
    for i in range(50):
        d.add_paragraph(f"Paragraph {i} alpha beta.")
    d.save(tmp_path / "a.docx")
    with zipfile.ZipFile(tmp_path / "a.docx") as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(tmp_path / "b.docx", "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "word/document.xml":
                data = data.replace(b"<w:p>", b'<w:p w:rsidR="00AB12CD">')
            z.writestr(info, data)
    doc = rdocx.Document.open(tmp_path / "b.docx")
    assert doc.try_replace_text("Paragraph 7 ", "Paragraph seven ") == 1
    doc.save(tmp_path / "c.docx")
    a, b = part(tmp_path / "b.docx", "word/document.xml"), part(tmp_path / "c.docx", "word/document.xml")
    assert a.count(b"xmlns:w=") == 1 and a.count(b"\n") == 1
    assert b.count(b"xmlns:w=") == 1 and b.count(b"\n") <= a.count(b"\n")  # no indentation
    assert abs(len(b) - len(a)) < 64


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
