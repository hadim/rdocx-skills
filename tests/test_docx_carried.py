"""docx: fixes and features the pinned integration build (integration/open-prs-2026-10-21) carries and upstream
`main` does not have yet, which the skills rely on. Each test names the pull request that carries it: delete the test
once upstream main has it. Each test builds its own input."""
import time
import zipfile

import pytest
import rdocx

from builders import bare_cell_docx, many_tables_docx, word_textbox_docx

WALK_BOUND_S = 20.0  # 0.1 s on the pinned build; #399 took its synthetic document's cell walk from hours to 0.2 s


# carried by tensorbee/rdocx#399, delete once upstream main has it
def test_walking_every_paragraph_and_cell_handle_stays_linear(tmp_path):
    """A walk over every body paragraph and every cell paragraph of 40 tables reads through handles without
    rescanning the document per handle (python-api.md, Handles)."""
    doc = rdocx.Document.open(many_tables_docx(tmp_path / "many.docx"))
    deadline = time.monotonic() + WALK_BOUND_S
    texts = []

    def read(paragraphs):
        for p in paragraphs:
            texts.append(p.text)
            if time.monotonic() > deadline:  # fail fast rather than hang for hours on a build without the fix
                pytest.fail(f"walking the handles took over {WALK_BOUND_S} s ({len(texts)} paragraphs read)")

    read(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                read(cell.paragraphs)
    assert len(texts) == 40 * 10 + 40 * 25
    assert texts[0] == "Body 0.0" and texts[-1] == "Cell 39.4.4"


# carried by tensorbee/rdocx#393, delete once upstream main has it
@pytest.mark.parametrize("edit", ["shading", "width"])
def test_cell_runs_stay_readable_after_a_cell_property_edit(tmp_path, edit):
    """On a cell without w:tcPr, setting its shading or width creates one: the cell's runs, fresh or held, still
    read (SKILL.md: formatting setters keep handles)."""
    doc = rdocx.Document.open(bare_cell_docx(tmp_path / "bare.docx"))
    cell = doc.tables[0].cell(0, 0)
    assert "tcPr" not in cell.xml
    held = cell.paragraphs[0].runs[0]
    if edit == "shading":
        cell.shading = "FFFF00"
    else:
        cell.width = rdocx.Inches(2)
    assert "tcPr" in doc.tables[0].cell(0, 0).xml
    assert [r.text for r in doc.tables[0].cell(0, 0).paragraphs[0].runs] == ["alpha"]
    assert held.text == "alpha"


# carried by tensorbee/rdocx#397, delete once upstream main has it
def test_remove_content_range_removes_blocks_in_one_edit():
    """remove_content_range(start, stop) removes body blocks start to stop (excluded), returns the count, takes a
    comment the range covers whole with it, and refuses a range outside the body."""
    doc = rdocx.Document()
    for i in range(6):
        doc.add_paragraph(f"p{i}")
    doc.add_comment(rdocx.RunRange(start=rdocx.RunPosition(body_index=2, run_index=0),
                                   end=rdocx.RunPosition(body_index=2, run_index=1)), author="Reviewer", text="Check.")
    assert doc.remove_content_range(1, 4) == 3
    assert [p.text for p in doc.paragraphs] == ["p0", "p4", "p5"]
    assert len(doc.comments) == 0
    assert doc.remove_content_range(1, 1) == 0
    with pytest.raises(IndexError):
        doc.remove_content_range(2, 9)
    assert [p.text for p in doc.paragraphs] == ["p0", "p4", "p5"]


# carried by tensorbee/rdocx#395, delete once upstream main has it
def test_default_paragraph_format_round_trips_and_shows_in_the_part_xml(tmp_path):
    """set_default_paragraph_format writes w:pPrDefault, keeps what it is not given, survives a save, and
    part_xml shows it before the save."""
    doc = rdocx.Document()
    doc.set_default_paragraph_format(space_before=rdocx.Pt(3), space_after=rdocx.Pt(6), line_spacing=1.15)
    doc.set_default_paragraph_format(space_after=rdocx.Pt(10))
    assert "/word/styles.xml" in doc.part_names
    assert "w:pPrDefault" in doc.part_xml("word/styles.xml")
    doc.save(tmp_path / "d.docx")
    fmt = rdocx.Document.open(tmp_path / "d.docx").default_paragraph_format
    assert isinstance(fmt, rdocx.DefaultParagraphFormat)
    assert (fmt.space_before, fmt.space_after, fmt.line_spacing) == (rdocx.Pt(3), rdocx.Pt(10), 1.15)


# carried by tensorbee/rdocx#395, delete once upstream main has it
def test_text_box_paragraph_handle(tmp_path):
    """StoryItem.paragraph gives a Paragraph handle on a text box paragraph; a replacement through it edits both
    copies Word writes (DrawingML and the VML fallback)."""
    doc = rdocx.Document.open(word_textbox_docx(tmp_path / "box.docx", footnote=False))
    item = next(it for it in doc.story_items if it.story.kind == "text_box" and it.kind == "paragraph")
    paragraph = item.paragraph
    assert paragraph.text == "Box NEEDLE"
    assert paragraph.replace_text("NEEDLE", "PIN", expect=1) == 1
    doc.save(tmp_path / "out.docx")
    with zipfile.ZipFile(tmp_path / "out.docx") as z:
        xml = z.read("word/document.xml").decode()
    assert xml.count("Box PIN") == 2 and "NEEDLE" not in xml


def _toc_document(instruction, outline_paragraphs=()):
    """Chapter 1 and 2 (Heading 1), each with two Heading 2 sections, then plain paragraphs carrying a direct
    outline level, and a TOC field at the top whose instruction is `instruction`."""
    doc = rdocx.Document()
    for i in (1, 2):
        doc.add_heading(f"Chapter {i}", level=1)
        for j in (1, 2):
            doc.add_heading(f"Section {i}.{j}", level=2)
            doc.add_paragraph("Body text.")
    for text, level in outline_paragraphs:
        p = doc.add_paragraph(text)
        p.replace_xml(p.xml.replace("<w:r>", f'<w:pPr><w:outlineLvl w:val="{level - 1}"/></w:pPr><w:r>', 1))
    doc.insert_toc(0, max_level=1)
    [field] = [p for p in doc.paragraphs if 'TOC \\o "1-1" \\h' in p.xml]
    field.replace_xml(field.xml.replace('TOC \\o "1-1" \\h', instruction))
    return doc


def _toc_entries(doc):
    rep = doc.rebuild_toc()
    assert list(rep.diagnostics) == []
    return [p.text.rsplit("\t", 1)[0].strip() for p in doc.paragraphs if (p.style or "").upper().startswith("TOC")]


# carried by tensorbee/rdocx#401, delete once upstream main has it
@pytest.mark.parametrize("instruction", ['TOC \\o "1-1" \\h \\z \\u', 'TOC \\o "1-1" \\u'])
def test_toc_depth_holds_with_outline_switch(instruction):
    """`\\o "1-N"` lists heading levels 1 to N only, with `\\u` as without it: upstream main (2a88a8f5, F-283)
    lets the Heading 2 entries in under `\\o "1-1" \\u` (tensorbee/rdocx#400); Word lists Heading 1 only."""
    assert _toc_entries(_toc_document(instruction)) == ["Chapter 1", "Chapter 2"]


# carried by tensorbee/rdocx#401, delete once upstream main has it
def test_toc_outline_switch_stays_within_depth():
    """`\\u` adds the paragraphs whose own outline level lies within `\\o "1-N"`, never beyond: Word for Mac
    lists the level-1 paragraph and leaves out the level-2 one; upstream main lists both (tensorbee/rdocx#400)."""
    doc = _toc_document('TOC \\o "1-1" \\h \\z \\u', [("Outline two", 2), ("Outline one", 1)])
    assert _toc_entries(doc) == ["Chapter 1", "Chapter 2", "Outline one"]
