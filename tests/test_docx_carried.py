"""docx: what the pinned integration build carries and upstream `main` does not have yet (CLAUDE.md, Rules). Each
test names the pull request that carries it; delete it once upstream `main` or a release has that behaviour. Each
test builds its own input."""
import re
import sys
from pathlib import Path

import pytest
import rdocx

from builders import body_docx, resolves_both_ways
from conftest import run

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "plugin" / "skills" / "docx" / "scripts"))
import docx_ops  # noqa: E402

STAMP = "2026-11-09T12:00:00Z"


# ---- compare -----------------------------------------------------------------------------------------------------

def word_redline(original_runs, edited_runs):
    """The word-granularity redline of a paragraph made of original_runs against one made of edited_runs:
    (the redlined document, deleted texts, inserted texts)."""
    original, edited = rdocx.Document(), rdocx.Document()
    p = original.add_paragraph(original_runs[0])
    for text in original_runs[1:]:
        p.add_run(text)
    q = edited.add_paragraph(edited_runs[0])
    for text in edited_runs[1:]:
        q.add_run(text)
    out = rdocx.Document.from_bytes(original.to_bytes())
    out.compare(edited, "Reviewer", STAMP, granularity="word")
    xml = out.part_xml("word/document.xml")
    deleted = re.findall(r"<w:delText[^>]*>([^<]*)</w:delText>", xml)
    inserted = [re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<w:ins\b.*?</w:ins>", xml, re.S)]
    return out, deleted, inserted


# carried: tensorbee/rdocx#450 (issue #446 case 1)
@pytest.mark.parametrize("edited", [["Pay within sixty days."], ["Pay wi", "thin sixty days."]],
                         ids=["one run", "split runs"])
def test_compare_by_word_keeps_words_whole(edited):
    """At word granularity, thirty -> sixty gives one deletion ("thirty") and one insertion ("sixty"), and the
    unchanged words of the paragraph stay in runs of whole words (python-api.md, compare; v0.16.4 wrote 8 runs
    here,
    0.17.0 one run per character)."""
    out, deleted, inserted = word_redline(["Pay within thirty days."], edited)
    assert (deleted, inserted) == (["thirty"], ["sixty"])
    runs = out.paragraphs[0].runs
    assert len(runs) <= 2 * len("Pay within thirty days.".split()), [r.text for r in runs]


# ---- comments ----------------------------------------------------------------------------------------------------

def thread_document():
    """Two paragraphs; a comment on "about 12 mm" in the second and a reply to it: (document, root id, reply id)."""
    doc = rdocx.Document()
    doc.add_paragraph("Keep the bearing positions here.")
    doc.add_paragraph("The bearing at pier 2 moved about 12 mm.")
    cid = docx_ops.comment_on_text(doc, "about 12 mm", "Which reference?", "Reviewer", date=STAMP)
    rid = doc.reply_to(cid, author="Author", text="The 2019 survey marks.", date=STAMP)
    return doc, cid, rid


# carried: tensorbee/rdocx#448 (issue #446 case 2)
def test_moving_a_thread_moves_its_replies(tmp_path, rdocx_cli):
    """move_comment_to_text and `rdocx comment move` move a thread with its replies (python-api.md, comments):
    removing the paragraph the thread left keeps every comment. In 0.17.0 the move took the root only."""
    doc, cid, rid = thread_document()
    src = tmp_path / "thread.docx"
    doc.save(str(src))
    res = run([rdocx_cli, "comment", "move", src, str(cid), "--text", "bearing positions", "-o", tmp_path / "cli.docx"])
    assert res.returncode == 0, res.stderr
    moved = rdocx.Document.open(str(tmp_path / "cli.docx"))
    assert {c.id: c.anchor_text for c in moved.comments} == {cid: "bearing positions", rid: "bearing positions"}
    doc.move_comment_to_text(cid, "bearing positions")
    assert {c.id: c.anchor_text for c in doc.comments} == {cid: "bearing positions", rid: "bearing positions"}
    doc.remove_content(doc.find_content_index(doc.paragraphs[1]))
    assert [(c.id, c.parent_id) for c in doc.comments] == [(cid, None), (rid, cid)]


# ---- tables ------------------------------------------------------------------------------------------------------

# carried: tensorbee/rdocx#448 (issue #446 case 3)
def test_grid_span_consumes_an_empty_cell_holding_paragraph_properties():
    """set_cell_grid_span consumes an empty cell whose paragraph carries properties (a right-aligned totals column),
    as Cell.merge does (python-api.md, Tables) (0.17.0 raised "horizontal merge cannot
    consume a nonempty cell")."""
    doc = rdocx.Document()
    table = doc.add_table(1, 3, style="Table Grid")
    table.cell(0, 0).text = "Total"
    table.cell(0, 1).paragraphs[0].alignment = rdocx.WD_ALIGN_PARAGRAPH.RIGHT
    doc.tables[0].set_cell_grid_span(0, 0, 2)
    assert [p.text for p in doc.tables[0].cell(0, 0).paragraphs] == ["Total"]


# ---- compare: paragraph boundaries -------------------------------------------------------------------------------

def para(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def linked(before, text):
    return (f'<w:p><w:r><w:t xml:space="preserve">{before}</w:t></w:r><w:hyperlink w:anchor="target">'
            f"<w:r><w:t>{text}</w:t></w:r></w:hyperlink></w:p>")


def in_cell(*paragraphs):
    cell = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{}</w:tc>'
    return (para("Lead.") + '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
            f'<w:tr>{cell.format("".join(paragraphs))}{cell.format(para("x"))}</w:tr></w:tbl>' + para("Tail."))


HEADING = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc1"/>'
           '<w:r><w:t>New heading</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>')


# carried: tensorbee/rdocx#454 (issue #453)
@pytest.mark.parametrize("original, edited", [
    (in_cell(para("Open the window,"), para("then wait a while."), para("Close the door.")),
     in_cell(para("Open the window."), para("Close the door."))),
    (in_cell(para("One two three."), para("Four.")), in_cell(para("One two"), para("three."), para("Four."))),
    (in_cell(para("Call the desk.")), in_cell(linked("Write to ", "the desk"))),
    (para("Lead.") + para("Call the desk."), para("Lead.") + linked("Write to ", "the desk")),
    (para("Lead.") + linked("See ", "the site") + para("Old text.") + para("Tail."),
     para("Lead.") + linked("Look at ", "the site") + HEADING + para("Tail.")),
], ids=["merge in a cell", "split in a cell", "one-paragraph cell gains a link", "last paragraph gains a link",
        "paragraph becomes a bookmarked heading beside a link"])
def test_compare_tracks_paragraph_boundary_changes(tmp_path, original, edited):
    """Paragraphs of a table cell that merge or split, a final paragraph that gains a hyperlink, and a paragraph
    that becomes a bookmarked heading beside a hyperlink compare at word granularity, and the redline resolves to
    each version (python-api.md, compare). 0.17.0 refused each ("cannot revise cell boundary structures", "cannot
    remove the final paragraph mark", "cannot revise paragraph boundary structures")."""
    resolves_both_ways(body_docx(tmp_path / "a.docx", original), body_docx(tmp_path / "b.docx", edited))
