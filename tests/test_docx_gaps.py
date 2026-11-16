"""docx: known gaps of the pinned upstream release (strict xfails, tests/gaps.py). Each test names its gap and,
once reported, its upstream issue; it XPASSes the day a pin closes the gap. Each test builds its own input."""
import re
import sys
import time
from pathlib import Path

import pytest
import rdocx

from builders import body_docx, heading_body, resolves_both_ways
from conftest import run

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "plugin" / "skills" / "docx" / "scripts"))
import docx_ops  # noqa: E402

STAMP = "2026-11-09T12:00:00Z"


# ---- compare -----------------------------------------------------------------------------------------------------

# upstream: tensorbee/rdocx#453
@pytest.mark.gap("docx-compare-renumbered-bookmarks")
def test_compare_versions_whose_heading_bookmarks_were_renumbered(tmp_path):
    """Two versions whose producer named the heading bookmarks _Toc1, _Toc2, ... in document order, a heading
    inserted before the others: compare gives a redline that resolves both ways (gaps.md). rdocx refuses, because
    the names moved to other headings and the redline would hold them twice; docx_ops.strip_generated_bookmarks on
    both copies is the workaround (test_skill_scripts.py)."""
    original = body_docx(tmp_path / "a.docx", heading_body([("Alpha", 1, "_Toc1", "First body."),
                                                            ("Beta", 2, "_Toc2", "Second body.")]))
    edited = body_docx(tmp_path / "b.docx", heading_body([("New", 1, "_Toc1", "New body."),
                                                          ("Alpha", 2, "_Toc2", "First body."),
                                                          ("Beta", 3, "_Toc3", "Second body.")]))
    resolves_both_ways(original, edited)


# upstream: tensorbee/rdocx#430
@pytest.mark.gap("docx-compare-moved-footnote-reference")
def test_compare_a_new_footnote_before_an_existing_one(tmp_path):
    """The edited version adds a footnote before an existing one, so the existing note's id now sits on another
    reference: compare gives a redline that resolves both ways (gaps.md). rdocx refuses such a comparison
    atomically, where 0.17.0 wrote a result Word could not open."""
    original, edited = rdocx.Document(), rdocx.Document()
    original.add_paragraph("Lead.")
    original.add_footnote(original.add_paragraph("Alpha."), "The first note.")
    edited.add_footnote(edited.add_paragraph("Lead."), "A new note.")
    edited.add_footnote(edited.add_paragraph("Alpha."), "The first note.")
    original.save(str(tmp_path / "a.docx"))
    edited.save(str(tmp_path / "b.docx"))
    resolves_both_ways(tmp_path / "a.docx", tmp_path / "b.docx")


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
