"""The helper scripts shipped in the skills must keep working on the pinned build: they encode the
workarounds the skills teach, so a pin that changes an underlying behaviour shows up here first."""
import os
import subprocess
import sys
from pathlib import Path

import pytest
import rdocx

from builders import every_story_docx, word_textbox_docx, wrapped_run_docx, wrapped_text_docx

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "skills" / "docx" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "pptx" / "scripts"))
import docx_ops  # noqa: E402
import pptx_ops  # noqa: E402
from conftest import BIN, digest  # noqa: E402


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
    when rdocx's count matches (a Word text box counts twice while a footnote is skipped)."""
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
    os.chmod(src, 0o644)
    docx_ops.save_atomic(rdocx.Document.open(src), tmp_path / "o.docx", src)
    assert os.stat(tmp_path / "o.docx").st_mode & 0o777 == 0o644


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
def test_comment_on_text_anchors_exactly_or_refuses(tmp_path, wrapper, anchor):
    """The helper's contract: the comment lands on exactly the anchor, or EditError and doc unchanged (today the
    runs after a content control or a tracked insertion are refused: gap comment-runposition-sdt)."""
    doc = rdocx.Document.open(wrapped_run_docx(tmp_path / "w.docx", wrapper))
    before = doc.to_bytes()
    try:
        cid = docx_ops.comment_on_text(doc, anchor, "x", "Reviewer")
    except docx_ops.EditError as e:
        assert "comment-runposition-sdt" in str(e) and digest(doc.to_bytes()) == digest(before)
    else:
        assert anchored(doc.to_bytes(), cid) == anchor


def test_comment_on_text_in_the_report_content_control(report_docx):
    doc = rdocx.Document.open(report_docx)
    before = doc.to_bytes()
    try:
        cid = docx_ops.comment_on_text(doc, "74 out of 100", "x", "Reviewer")
    except docx_ops.EditError:
        assert digest(doc.to_bytes()) == digest(before)
    else:
        assert anchored(doc.to_bytes(), cid) == "74 out of 100"


def test_command_line_entry_points(report_docx, tmp_path):
    py = str(BIN / "python")  # the runtime Python an agent gets
    script = ROOT / "skills" / "docx" / "scripts" / "docx_ops.py"
    res = subprocess.run([py, script, "replace", report_docx, tmp_path / "o.docx", "--edit", "footbridge", "bridge", "1"],
                         capture_output=True, text=True)
    assert res.returncode == 1 and "refused" in res.stderr and not (tmp_path / "o.docx").exists()
    res = subprocess.run([py, script, "pages", report_docx], capture_output=True, text=True, check=True)
    assert int(res.stdout) == docx_ops.pages(report_docx)
    res = subprocess.run([py, script, "comment", report_docx, tmp_path / "c.docx", "--anchor", "three-span steel and timber",
                          "--text", "x", "--author", "R", "--date", "2026-09-27T12:00:00Z"], capture_output=True, text=True, check=True)
    assert rdocx.Document.open(tmp_path / "c.docx").comments[0].date == "2026-09-27T12:00:00Z"
    path, _ = every_story_docx(tmp_path / "s.docx")
    res = subprocess.run([py, script, "count", path, "NEEDLE"], capture_output=True, text=True, check=True)
    assert '"footnote": 1' in res.stdout
    res = subprocess.run([py, script, "text", path], capture_output=True, text=True, check=True)
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
    assert os.stat(out).st_mode & 0o777 == 0o644


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
def test_count_reads_text_that_paragraph_text_misses(tmp_path, wrapper):
    path = wrapped_text_docx(tmp_path / "w.docx", wrapper)
    assert docx_ops.all_text(path) == [("body", "before MID after")]
