"""docx, review features: comments and threads, tracked changes, redline by comparison."""
import json
import os
import re
import shutil
import zipfile

import docx
import pytest
import rdocx

from builders import wrapped_run_docx
from conftest import STAMP, part, run


def two_paragraphs(path, first="Alpha paragraph with some words here.", second="Beta paragraph."):
    d = docx.Document()
    d.add_paragraph(first)
    d.add_paragraph(second)
    d.save(path)
    return path


def anchored(path, cid):
    x = part(path, "word/document.xml").decode()
    m = re.search(r'<w:commentRangeStart w:id="%d"/>(.*?)<w:commentRangeEnd w:id="%d"/>' % (cid, cid), x, re.S)
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", m.group(1))) if m else None


def run_range(bi, first, last):
    return rdocx.RunRange(start=rdocx.RunPosition(body_index=bi, run_index=first),
                          end=rdocx.RunPosition(body_index=bi, run_index=last))


def noop_saved(src, dst):
    """The original passed through rdocx once (open, save): the workaround for empty-comments-reserialised."""
    rdocx.Document.open(src).save(dst)
    return dst


def compare(cli, a, b, out):
    if os.path.exists(out):
        os.remove(out)
    return run([cli, "compare", a, b, "--author", "Reviewer", "--timestamp", STAMP, "-o", out])


# ---------------------------------------------------------------- comments
def test_comment_thread_python(tmp_path):
    doc = rdocx.Document.open(two_paragraphs(tmp_path / "c.docx"))
    rng = rdocx.RunRange(start=rdocx.RunPosition(body_index=0, run_index=0),
                         end=rdocx.RunPosition(body_index=0, run_index=1))
    cid = doc.add_comment(rng, author="Reviewer", text="Check this.", initials="R", date=STAMP)
    rid = doc.reply_to(cid, author="Author", text="Done.")
    doc.resolve_comment(cid)
    got = [(c.id, c.author, c.text, c.parent_id, c.resolved) for c in doc.comments]
    assert got == [(cid, "Reviewer", "Check this.", None, True), (rid, "Author", "Done.", cid, False)]
    doc.save(tmp_path / "c2.docx")
    assert [(c.author, c.text) for c in docx.Document(tmp_path / "c2.docx").comments] == [("Reviewer", "Check this."), ("Author", "Done.")]
    again = rdocx.Document.open(tmp_path / "c2.docx")
    again.remove_comment(cid)
    assert len(again.comments) == 0


def test_comment_thread_cli(rdocx_cli, tmp_path):
    src = two_paragraphs(tmp_path / "c.docx")
    a, b, c = tmp_path / "a.docx", tmp_path / "b.docx", tmp_path / "d.docx"
    add = run([rdocx_cli, "comment", "add", src, "--start-paragraph", "0", "--start-run", "0", "--end-paragraph", "0",
               "--end-run", "1", "--author", "Reviewer", "--text", "Why?", "-o", a, "--json"], check=True)
    cid = json.loads(add.stdout)
    run([rdocx_cli, "comment", "reply", a, "--id", "0", "--author", "Author", "--text", "Because.", "-o", b], check=True)
    run([rdocx_cli, "comment", "resolve", b, "--id", "0", "-o", c], check=True)
    listed = json.loads(run([rdocx_cli, "comment", "list", "--json", c], check=True).stdout)["comments"]
    assert [x["text"] for x in listed] == ["Why?", "Because."] and listed[0]["resolved"] is True
    assert cid  # the operation record is JSON


def test_comment_ids_are_stable_across_save(tmp_path):
    doc = rdocx.Document.open(two_paragraphs(tmp_path / "c.docx"))
    a = doc.add_comment(run_range(1, 0, 1), author="A", text="on the second paragraph")
    b = doc.add_comment(run_range(0, 0, 1), author="A", text="on the first paragraph")
    doc.save(tmp_path / "d.docx")
    assert [(c.id, c.text) for c in rdocx.Document.open(tmp_path / "d.docx").comments] == [
        (a, "on the second paragraph"), (b, "on the first paragraph")]


def test_comment_date_is_written_only_when_given(tmp_path):
    doc = rdocx.Document.open(two_paragraphs(tmp_path / "c.docx"))
    doc.add_comment(run_range(0, 0, 1), author="A", text="undated")
    doc.add_comment(run_range(1, 0, 1), author="A", text="dated", date=STAMP)
    doc.save(tmp_path / "d.docx")
    assert [(c.text, c.date) for c in rdocx.Document.open(tmp_path / "d.docx").comments] == [("undated", None), ("dated", STAMP)]


@pytest.mark.gap("comment-date-cli")
def test_cli_comment_add_takes_a_date(rdocx_cli, tmp_path):
    src = two_paragraphs(tmp_path / "c.docx")
    res = run([rdocx_cli, "comment", "add", src, "--start-paragraph", "0", "--start-run", "0", "--end-paragraph", "0",
               "--end-run", "1", "--author", "R", "--text", "x", "--date", STAMP, "-o", tmp_path / "d.docx"])
    assert res.returncode == 0 and rdocx.Document.open(tmp_path / "d.docx").comments[0].date == STAMP


def test_comment_on_a_table_cell_through_a_story_range(tmp_path):
    d = docx.Document()
    d.add_paragraph("before")
    d.add_table(rows=1, cols=2).cell(0, 1).text = "cell text"
    d.save(tmp_path / "c.docx")
    doc = rdocx.Document.open(tmp_path / "c.docx")
    item = next(it for it in doc.story_items if it.kind == "paragraph" and it.text == "cell text")
    rng = rdocx.StoryRunRange(start=rdocx.StoryRunPosition(item=item, run_index=0),
                              end=rdocx.StoryRunPosition(item=item, run_index=1))
    cid = doc.add_comment(rng, author="A", text="on a cell", date=STAMP)
    doc.save(tmp_path / "d.docx")
    assert anchored(tmp_path / "d.docx", cid) == "cell text"


@pytest.mark.gap("comment-runposition-sdt")
@pytest.mark.parametrize("wrapper", ["sdt", "ins"])
def test_comment_run_index_counts_wrapped_runs(tmp_path, wrapper):
    path = wrapped_run_docx(tmp_path / "w.docx", wrapper)
    doc = rdocx.Document.open(path)
    assert [r.text for r in doc.paragraphs[0].runs] == ["before ", "TARGET", " after"]
    cid = doc.add_comment(run_range(0, 1, 2), author="A", text="x")
    doc.save(tmp_path / "c.docx")
    assert anchored(tmp_path / "c.docx", cid) == "TARGET"


@pytest.mark.gap("comment-runposition-sdt")
def test_cli_comment_run_index_counts_wrapped_runs(rdocx_cli, tmp_path):
    path = wrapped_run_docx(tmp_path / "w.docx", "sdt")
    run([rdocx_cli, "comment", "add", path, "--start-paragraph", "0", "--start-run", "1", "--end-paragraph", "0",
         "--end-run", "2", "--author", "A", "--text", "x", "-o", tmp_path / "c.docx"], check=True)
    assert anchored(tmp_path / "c.docx", 0) == "TARGET"


def test_comment_cli_refuses_to_overwrite_its_input(rdocx_cli, tmp_path):
    src = two_paragraphs(tmp_path / "c.docx")
    before = src.read_bytes()
    res = run([rdocx_cli, "comment", "add", src, "--start-paragraph", "0", "--start-run", "0", "--end-paragraph", "0",
               "--end-run", "1", "--author", "R", "--text", "x", "-o", src])
    assert res.returncode != 0 and src.read_bytes() == before


# ---------------------------------------------------------------- tracked changes
def test_compare_then_accept_or_reject(rdocx_cli, tmp_path):
    a = two_paragraphs(tmp_path / "a.docx")
    b = two_paragraphs(tmp_path / "b.docx", first="Alpha paragraph with other words here.")
    red = tmp_path / "red.docx"
    assert compare(rdocx_cli, a, b, red).returncode == 0
    doc = rdocx.Document.open(red)
    kinds = sorted({r.kind for r in doc.revisions})
    assert "insertion" in kinds and "deletion" in kinds
    acc = rdocx.Document.open(red)
    acc.accept_all()
    rej = rdocx.Document.open(red)
    rej.reject_all()
    assert acc.paragraphs[0].text == "Alpha paragraph with other words here."
    assert rej.paragraphs[0].text == "Alpha paragraph with some words here."


def test_revision_cli_list_and_accept_by_author(rdocx_cli, tmp_path):
    a = two_paragraphs(tmp_path / "a.docx")
    b = two_paragraphs(tmp_path / "b.docx", second="Beta paragraph, edited.")
    red, acc = tmp_path / "red.docx", tmp_path / "acc.docx"
    compare(rdocx_cli, a, b, red)
    listed = json.loads(run([rdocx_cli, "revision", "list", "--json", red], check=True).stdout)["revisions"]
    assert listed and all(r["author"] == "Reviewer" for r in listed)
    run([rdocx_cli, "revision", "accept", red, "--author", "Reviewer", "-o", acc], check=True)
    assert json.loads(run([rdocx_cli, "revision", "list", "--json", acc], check=True).stdout)["revisions"] == []


def test_accept_revisions_in_a_date_range(rdocx_cli, tmp_path):
    a = two_paragraphs(tmp_path / "a.docx")
    b = two_paragraphs(tmp_path / "b.docx", first="Alpha paragraph with other words here.")
    compare(rdocx_cli, a, b, tmp_path / "red.docx")
    doc = rdocx.Document.open(tmp_path / "red.docx")
    assert doc.accept_revisions_in_date_range(start="2026-09-27T00:00:00Z", end="2026-09-28T00:00:00Z") >= 2
    assert doc.paragraphs[0].text == "Alpha paragraph with other words here." and len(doc.revisions) == 0


def test_python_compare_method(tmp_path):
    a = rdocx.Document.open(two_paragraphs(tmp_path / "a.docx"))
    b = rdocx.Document.open(two_paragraphs(tmp_path / "b.docx", second="Gamma paragraph."))
    a.compare(b, "Reviewer", STAMP)
    assert len(a.revisions) >= 2


@pytest.mark.gap("compare-granularity")
def test_compare_marks_only_the_changed_word(rdocx_cli, tmp_path):
    text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor."
    a = two_paragraphs(tmp_path / "a.docx", first=text)
    b = two_paragraphs(tmp_path / "b.docx", first=text.replace("dolor", "DOLOR"))
    compare(rdocx_cli, a, b, tmp_path / "red.docx")
    deleted = "".join(re.findall(r"<w:delText[^>]*>([^<]*)</w:delText>", part(tmp_path / "red.docx", "word/document.xml").decode()))
    assert deleted.strip() == "dolor"


@pytest.mark.gap("compare-comments")
def test_compare_accepts_a_pair_whose_comments_differ(rdocx_cli, tmp_path):
    a = two_paragraphs(tmp_path / "a.docx")
    b = tmp_path / "b.docx"
    run([rdocx_cli, "comment", "add", a, "--start-paragraph", "1", "--start-run", "0", "--end-paragraph", "1",
         "--end-run", "1", "--author", "R", "--text", "New comment", "-o", b], check=True)
    assert compare(rdocx_cli, a, b, tmp_path / "red.docx").returncode == 0


@pytest.mark.gap("compare-sdt-id")
def test_compare_ignores_a_content_control_id(rdocx_cli, report_docx, tmp_path):
    b = tmp_path / "b.docx"
    with zipfile.ZipFile(report_docx) as zin, zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = data.replace(b'w:val="1209466531"', b'w:val="1209466532"')
            zout.writestr(info, data)
    assert compare(rdocx_cli, report_docx, b, tmp_path / "red.docx").returncode == 0


@pytest.mark.gap("empty-comments-reserialised")
def test_compare_against_its_own_rdocx_save(rdocx_cli, report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    assert compare(rdocx_cli, report_docx, tmp_path / "e.docx", tmp_path / "red.docx").returncode == 0


@pytest.mark.gap("compare-own-save-noise")
def test_compare_after_rdocx_edit_reports_only_the_edit(rdocx_cli, tmp_path):
    d = docx.Document()
    for t in ("Paragraph 1, lorem ipsum.", "WORD", "Paragraph 3, lorem ipsum."):
        d.add_paragraph(t)
    sect = d.sections[0]._sectPr
    pgsz = sect.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz")
    pgsz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}orient", "portrait")
    d.save(tmp_path / "a.docx")
    doc = rdocx.Document.open(tmp_path / "a.docx")
    doc.try_replace_text("WORD", "TERM")
    doc.save(tmp_path / "b.docx")
    compare(rdocx_cli, tmp_path / "a.docx", tmp_path / "b.docx", tmp_path / "red.docx")
    kinds = [r.kind for r in rdocx.Document.open(tmp_path / "red.docx").revisions]
    assert "section_property_change" not in kinds and len(kinds) == 2


def test_compare_after_a_plain_edit_of_an_rdocx_saved_original(rdocx_cli, report_docx, tmp_path):
    """The base of the two gap tests below: without the field refresh or the TOC rebuild, the pair compares."""
    base = noop_saved(report_docx, tmp_path / "base.docx")
    doc = rdocx.Document.open(base)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    assert compare(rdocx_cli, base, tmp_path / "e.docx", tmp_path / "red.docx").returncode == 0


@pytest.mark.gap("compare-packed-fields")
def test_compare_after_refreshing_packed_page_fields(rdocx_cli, report_docx, tmp_path):
    base = noop_saved(report_docx, tmp_path / "base.docx")
    doc = rdocx.Document.open(base)
    doc.update_page_fields()
    doc.save(tmp_path / "f.docx")
    res = compare(rdocx_cli, base, tmp_path / "f.docx", tmp_path / "red.docx")
    assert res.returncode == 0, res.stderr


@pytest.mark.gap("compare-rebuilt-toc")
def test_compare_after_toc_rebuild(rdocx_cli, report_docx, tmp_path):
    base = noop_saved(report_docx, tmp_path / "base.docx")
    doc = rdocx.Document.open(base)
    doc.try_replace_text("described", "outlined")  # an edit rewrites the identity attributes, then the TOC rebuilds
    doc.rebuild_toc()
    doc.save(tmp_path / "t.docx")
    res = compare(rdocx_cli, base, tmp_path / "t.docx", tmp_path / "red.docx")
    assert res.returncode == 0, res.stderr


def test_revision_accept_record_counts_every_story(rdocx_cli, tmp_path):
    """The workaround for revisions-main-story: `revision accept --json` counts revisions in every story."""
    d = docx.Document()
    d.add_paragraph("Body.")
    d.sections[0].footer.paragraphs[0].text = "Footer lorem ipsum"
    d.save(tmp_path / "a.docx")
    d = docx.Document(tmp_path / "a.docx")
    d.sections[0].footer.paragraphs[0].text = "Footer lorem IPSUM"
    d.save(tmp_path / "b.docx")
    compare(rdocx_cli, tmp_path / "a.docx", tmp_path / "b.docx", tmp_path / "red.docx")
    rec = json.loads(run([rdocx_cli, "revision", "accept", tmp_path / "red.docx", "-o", tmp_path / "acc.docx", "--json"],
                         check=True).stdout)
    assert rec["resolved"] >= 1 and rec["scope"] == "all-supported-stories"


@pytest.mark.gap("revisions-main-story")
def test_revisions_listed_across_stories(rdocx_cli, tmp_path):
    d = docx.Document()
    d.add_paragraph("Body.")
    d.sections[0].footer.paragraphs[0].text = "Footer lorem ipsum"
    d.save(tmp_path / "a.docx")
    d = docx.Document(tmp_path / "a.docx")
    d.sections[0].footer.paragraphs[0].text = "Footer lorem IPSUM"
    d.save(tmp_path / "b.docx")
    compare(rdocx_cli, tmp_path / "a.docx", tmp_path / "b.docx", tmp_path / "red.docx")
    assert b"<w:ins " in part(tmp_path / "red.docx", "word/footer1.xml")
    assert len(rdocx.Document.open(tmp_path / "red.docx").revisions) > 0
