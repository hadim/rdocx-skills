"""The helper scripts shipped in the skills must keep working on the pinned build: they encode the
workarounds the skills teach, so a pin that changes an underlying behaviour shows up here first."""
import os
import subprocess
import sys
from pathlib import Path

import pytest
import rdocx

from builders import cell_text_docx, every_story_docx, word_textbox_docx, wrapped_run_docx, wrapped_text_docx

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "skills" / "docx" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "pptx" / "scripts"))
import docx_ops  # noqa: E402
import pptx_ops  # noqa: E402
from conftest import digest, runtime_python  # noqa: E402


anchored = docx_ops.anchored_text


def test_replace_batch_is_all_or_nothing(report_docx, tmp_path):
    out = tmp_path / "o.docx"
    with pytest.raises(docx_ops.EditError, match="footbridge"):
        docx_ops.replace_batch(report_docx, out, [("described", "outlined", 1), ("footbridge", "bridge", 1)])
    assert not out.exists()
    docx_ops.replace_batch(report_docx, out, [("described", "outlined", 1)])
    assert any("outlined" in p.text for p in rdocx.Document.open(out).paragraphs)


@pytest.mark.parametrize("build", ["every_story", "word_textbox"])
def test_replace_batch_never_leaves_occurrences_silently(tmp_path, build):
    """The helper's contract, whatever the pin reaches: text left after the replacement makes it refuse, even
    when rdocx's count matches."""
    path = every_story_docx(tmp_path / "s.docx")[0] if build == "every_story" else word_textbox_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    total = sum(docx_ops.count(doc, "NEEDLE").values())
    probe = rdocx.Document.from_bytes(doc.to_bytes())
    reached = probe.try_replace_text("NEEDLE", "X")
    left = sum(docx_ops.count(probe, "NEEDLE").values())
    out = tmp_path / "o.docx"
    if left:
        with pytest.raises(docx_ops.EditError, match=rf"{left} occurrence\(s\) left"):
            docx_ops.replace_batch(path, out, [("NEEDLE", "X", reached)])
        assert not out.exists()
    docx_ops.replace_batch(path, out, [("NEEDLE", "X", reached)], allow_unreached=True)
    assert sum(docx_ops.count(rdocx.Document.open(out), "NEEDLE").values()) == left
    assert left < total


def test_count_and_all_text(tmp_path):
    """What a reader sees: body (cells, text boxes, content controls, insertions), headers and footers per
    variant part (their tables included), footnotes, endnotes; deletions out; a Word text box once."""
    path, _ = every_story_docx(tmp_path / "s.docx")
    doc = rdocx.Document.open(path)
    assert docx_ops.count(doc, "NEEDLE") == {"body": 6, "header": 1, "footer": 3, "footnote": 1, "endnote": 1}
    body = [t for k, t in docx_ops.all_text(path) if k == "body" and "NEEDLE" in t]
    assert "Text box NEEDLE" in body and "Tracked: ins NEEDLE" in body and not any("del NEEDLE" in t for t in body)
    assert docx_ops.count(word_textbox_docx(tmp_path / "w.docx"), "NEEDLE") == {"body": 1, "footnote": 1}


def test_save_refuses_to_overwrite_the_input(report_docx):
    with pytest.raises(docx_ops.EditError):
        docx_ops.save_atomic(rdocx.Document.open(report_docx), report_docx, report_docx)


def test_save_keeps_the_input_file_mode(report_docx, copy_of, tmp_path):
    src = copy_of(report_docx)
    os.chmod(src, 0o644)  # on Windows, only the read-only bit of a mode exists: this one reads 0o666 there
    docx_ops.save_atomic(rdocx.Document.open(src), tmp_path / "o.docx", src)
    assert os.stat(tmp_path / "o.docx").st_mode & 0o777 == os.stat(src).st_mode & 0o777


def test_comment_on_text_after_tables(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    cid = docx_ops.comment_on_text(doc, "reviewed against the photographs", "Which photographs?", "Reviewer")
    docx_ops.save_atomic(doc, tmp_path / "c.docx", report_docx)
    assert anchored(tmp_path / "c.docx", cid) == "reviewed against the photographs"


def test_comment_on_text_mid_run_in_first_paragraphs(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    cid = docx_ops.comment_on_text(doc, "three-span steel and timber", "Check the span count.", "Reviewer")
    docx_ops.save_atomic(doc, tmp_path / "c.docx", report_docx)
    assert anchored(tmp_path / "c.docx", cid) == "three-span steel and timber"


def test_comment_on_text_dates_the_comment(report_docx):
    doc = rdocx.Document.open(report_docx)
    docx_ops.comment_on_text(doc, "three-span steel and timber", "x", "Reviewer", date="2026-09-27T12:00:00Z")
    assert doc.comments[0].date == "2026-09-27T12:00:00Z"
    docx_ops.comment_on_text(doc, "reviewed against the photographs", "y", "Reviewer")
    assert doc.comments[1].date  # now, by default


def test_comment_on_a_heading_skips_the_table_of_contents(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    cid = docx_ops.comment_on_text(doc, "Scope and method", "Rename?", "Reviewer")
    docx_ops.save_atomic(doc, tmp_path / "c.docx", report_docx)
    assert anchored(tmp_path / "c.docx", cid) == "Scope and method"


@pytest.mark.parametrize("wrapper", ["sdt", "ins"])
@pytest.mark.parametrize("anchor", ["before", "TARGET", "after"])
def test_comment_on_text_anchors_exactly_around_wrapped_runs(tmp_path, wrapper, anchor):
    """The comment lands on exactly the anchor, before, inside and after a content control or a tracked
    insertion."""
    doc = rdocx.Document.open(wrapped_run_docx(tmp_path / "w.docx", wrapper))
    cid = docx_ops.comment_on_text(doc, anchor, "x", "Reviewer")
    assert anchored(doc.to_bytes(), cid) == anchor


@pytest.mark.parametrize("wrapper", ["fldSimple", "smartTag", "customXml"])
def test_comment_on_text_counts_run_offsets_around_fields_and_smart_tags(tmp_path, wrapper):
    """Paragraph.text holds the text of a simple field, a smart tag or custom XML, Paragraph.runs does not:
    an anchor after one lands on exactly its text, an anchor inside one is refused with that reason."""
    doc = rdocx.Document.open(wrapped_text_docx(tmp_path / "w.docx", wrapper))
    assert doc.paragraphs[0].text == "before MID after"
    for anchor in ("before", "after"):
        cid = docx_ops.comment_on_text(doc, anchor, "x", "Reviewer")
        assert anchored(doc.to_bytes(), cid) == anchor
    i, start, _ = docx_ops.locate(doc, "after")
    first, last = docx_ops.isolate(doc, i, start, start + len("after"))
    assert [r.text for r in doc.paragraphs[i].runs[first:last]] == ["after"]
    for anchor in ("MID", "MID after"):
        with pytest.raises(docx_ops.EditError, match="simple field, a smart tag or a custom XML"):
            docx_ops.comment_on_text(doc, anchor, "x", "Reviewer")


def test_comment_on_text_in_the_report_content_control(report_docx):
    doc = rdocx.Document.open(report_docx)
    cid = docx_ops.comment_on_text(doc, "74 out of 100", "x", "Reviewer")
    assert anchored(doc.to_bytes(), cid) == "74 out of 100"


def test_comment_on_part_of_a_table_cell_paragraph(tmp_path):
    """in_tables=True counts the anchor in table cells only and anchors on exactly it, the cell's run split
    with its format kept."""
    doc = rdocx.Document.open(cell_text_docx(tmp_path / "t.docx"))
    cid = docx_ops.comment_on_text(doc, "beta", "x", "Reviewer", in_tables=True, date="2026-09-30T12:00:00Z")
    data = doc.to_bytes()
    assert anchored(data, cid) == "beta" and docx_ops.comment_landings(data)[cid] is True
    runs = [r for r in doc.tables[0].cell(0, 0).paragraphs[0].runs if r.text]
    assert [r.text for r in runs] == ["alpha ", "beta", " gamma"] and all(r.font.bold for r in runs)
    assert doc.comments[0].date == "2026-09-30T12:00:00Z"


def test_comment_on_text_in_tables_counts_nested_cells_and_refuses_past_the_last(tmp_path):
    doc = rdocx.Document.open(cell_text_docx(tmp_path / "t.docx"))
    cid = docx_ops.comment_on_text(doc, "beta", "x", "Reviewer", occurrence=2, in_tables=True)
    data = doc.to_bytes()
    assert anchored(data, cid) == "beta" and docx_ops.comment_landings(data)[cid] is True
    assert [r.text for r in doc.tables[0].cell(0, 0).paragraphs[0].runs] == ["alpha beta gamma"]  # not the first
    before = doc.to_bytes()
    with pytest.raises(docx_ops.EditError, match="not found in table cells"):
        docx_ops.comment_on_text(doc, "beta", "x", "Reviewer", occurrence=3, in_tables=True)
    assert digest(doc.to_bytes()) == digest(before)
    cid = docx_ops.comment_on_text(doc, "beta", "x", "Reviewer")  # the default still counts body paragraphs
    assert docx_ops.comment_landings(doc.to_bytes())[cid] is False


def test_command_line_entry_points(report_docx, tmp_path):
    py = runtime_python()  # the runtime Python an agent gets
    script = ROOT / "skills" / "docx" / "scripts" / "docx_ops.py"
    res = subprocess.run([*py, script, "replace", report_docx, tmp_path / "o.docx", "--edit", "footbridge", "bridge", "1"],
                         capture_output=True, text=True)
    assert res.returncode == 1 and "refused" in res.stderr and not (tmp_path / "o.docx").exists()
    res = subprocess.run([*py, script, "pages", report_docx], capture_output=True, text=True, check=True)
    assert int(res.stdout) == docx_ops.pages(report_docx)
    res = subprocess.run([*py, script, "comment", report_docx, tmp_path / "c.docx", "--anchor", "three-span steel and timber",
                          "--text", "x", "--author", "R", "--date", "2026-09-27T12:00:00Z"], capture_output=True, text=True, check=True)
    assert rdocx.Document.open(tmp_path / "c.docx").comments[0].date == "2026-09-27T12:00:00Z"
    cells = cell_text_docx(tmp_path / "cells.docx")
    subprocess.run([*py, script, "comment", cells, tmp_path / "cc.docx", "--anchor", "beta", "--in-tables", "--text", "x",
                    "--author", "R"], capture_output=True, text=True, check=True)
    assert docx_ops.comment_landings(tmp_path / "cc.docx") == {0: True} and anchored(tmp_path / "cc.docx", 0) == "beta"
    path, _ = every_story_docx(tmp_path / "s.docx")
    res = subprocess.run([*py, script, "count", path, "NEEDLE"], capture_output=True, text=True, check=True)
    assert '"footnote": 1' in res.stdout
    res = subprocess.run([*py, script, "text", path], capture_output=True, text=True, check=True)
    assert "footnote\tword/footnotes.xml\t\t Footnote NEEDLE." in res.stdout


def test_pptx_replace_batch(deck_pptx, tmp_path, copy_of):
    out = tmp_path / "o.pptx"
    with pytest.raises(pptx_ops.EditError):
        pptx_ops.replace_batch(deck_pptx, out, [("Riverton Footbridge", "Kestrel Footbridge", 1), ("option", "OPTION", 1)])
    assert not out.exists()
    src = copy_of(deck_pptx)
    os.chmod(src, 0o644)
    pptx_ops.replace_batch(src, out, [("Riverton Footbridge", "Kestrel Footbridge", 1)])
    assert "Kestrel Footbridge" in pptx_ops.cli("text", out).stdout
    assert os.stat(out).st_mode & 0o777 == os.stat(src).st_mode & 0o777


def test_pptx_overflow_report(deck_pptx):
    assert [(s, sid) for s, sid, _ in pptx_ops.overflowing(deck_pptx)] == [(4, 19)]


def test_toc_rebuild_workaround(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    assert docx_ops.rebuild_toc(doc).entry_count == 21
    docx_ops.save_atomic(doc, tmp_path / "t.docx", report_docx)
    assert [p.text for p in rdocx.Document.open(tmp_path / "t.docx").paragraphs if p.style == "Heading1"][0] == "Summary"


def test_image_helpers():
    png = docx_ops.solid_png(30, 20, (1, 2, 3))
    assert docx_ops.image_size(png) == (30, 20) and pptx_ops.image_size(png) == (30, 20)
    from PIL import Image
    import io
    assert Image.open(io.BytesIO(png)).convert("RGB").getpixel((5, 5)) == (1, 2, 3)
    jpg = io.BytesIO()
    Image.new("RGB", (41, 17)).save(jpg, "JPEG")
    assert docx_ops.image_size(jpg.getvalue()) == (41, 17)


def test_fix_template_content_type(tmp_path, report_docx):
    import zipfile
    with zipfile.ZipFile(report_docx) as src, zipfile.ZipFile(tmp_path / "t.docx", "w") as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"document.main+xml", b"template.main+xml")
            dst.writestr(info, data)
    assert docx_ops.fix_template_content_type(tmp_path / "t.docx") is True
    assert docx_ops.fix_template_content_type(tmp_path / "t.docx") is False
    with zipfile.ZipFile(tmp_path / "t.docx") as z, zipfile.ZipFile(report_docx) as r:
        assert z.namelist() == r.namelist() and z.read("word/document.xml") == r.read("word/document.xml")
    assert rdocx.Document.open(tmp_path / "t.docx").paragraphs


@pytest.mark.parametrize("wrapper", ["fldSimple", "smartTag", "customXml"])
def test_count_reads_text_inside_simple_fields_smart_tags_and_custom_xml(tmp_path, wrapper):
    path = wrapped_text_docx(tmp_path / "w.docx", wrapper)
    assert docx_ops.all_text(path) == [("body", "before MID after")]


def test_story_paragraphs_lists_a_paragraph_removed_by_a_tracked_deletion(tmp_path):
    a, b = rdocx.Document(), rdocx.Document()
    for text in ("Alpha", "Gone", "Omega"):
        a.add_paragraph(text)
    for text in ("Alpha", "Omega"):
        b.add_paragraph(text)
    a.compare(b, "Reviewer", "2026-09-27T12:00:00Z", granularity="word")
    a.save(tmp_path / "r.docx")
    texts = [t for _, kind, _, t in docx_ops.story_paragraphs(tmp_path / "r.docx") if kind == "body"]
    assert texts == ["Alpha", "", "Omega"]
