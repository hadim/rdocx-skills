"""docx: known gaps of the pinned upstream release (strict xfails, tests/gaps.py). Each test names its gap and,
once reported, its upstream issue; it XPASSes the day a pin closes the gap. Each test builds its own input."""
import re
import sys
import time
from pathlib import Path

import pytest
import rdocx

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


# upstream: tensorbee/rdocx#446 (case 1)
@pytest.mark.gap("docx-compare-word-splits-characters")
@pytest.mark.parametrize("edited", [["Pay within sixty days."], ["Pay wi", "thin sixty days."]],
                         ids=["one run", "split runs"])
def test_compare_by_word_keeps_words_whole(edited):
    """At word granularity, thirty -> sixty gives one deletion ("thirty") and one insertion ("sixty"), and the
    unchanged words of the paragraph stay in runs of whole words (python-api.md, compare; v0.16.4 wrote 8 runs
    here). rdocx 0.17.0 writes each character of the changed paragraph as its own run and each deleted or inserted
    letter as its own w:delText / w:t: about twice the XML and ten times the compare time on a long document."""
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


# upstream: tensorbee/rdocx#446 (case 2)
@pytest.mark.gap("docx-move-thread-leaves-replies")
def test_moving_a_thread_moves_its_replies(tmp_path, rdocx_cli):
    """move_comment_to_text and `rdocx comment move` move a thread with its replies (python-api.md, comments):
    removing the paragraph the thread left keeps every comment. Since replies got their own range (0.17.0) the move
    takes the root only, a reply cannot be moved ("not a movable root"), and removing the old paragraph drops the
    replies."""
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

# upstream: tensorbee/rdocx#446 (case 3)
@pytest.mark.gap("docx-grid-span-refuses-formatted-empty-cell")
def test_grid_span_consumes_an_empty_cell_holding_paragraph_properties():
    """set_cell_grid_span consumes an empty cell whose paragraph carries properties (a right-aligned totals column),
    as Cell.merge does (python-api.md, Tables): it raises "horizontal merge cannot consume a nonempty cell"."""
    doc = rdocx.Document()
    table = doc.add_table(1, 3, style="Table Grid")
    table.cell(0, 0).text = "Total"
    table.cell(0, 1).paragraphs[0].alignment = rdocx.WD_ALIGN_PARAGRAPH.RIGHT
    doc.tables[0].set_cell_grid_span(0, 0, 2)
    assert [p.text for p in doc.tables[0].cell(0, 0).paragraphs] == ["Total"]


# ---- edit costs --------------------------------------------------------------------------------------------------

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
