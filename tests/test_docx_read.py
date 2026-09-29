"""docx, reading: text, structure, JSON views, metadata, hyperlinks, sections, styles."""
import json
import zipfile

import docx
import pytest
import rdocx

from builders import every_story_docx
from conftest import run


def test_cli_text_plain(rdocx_cli, report_docx):
    out = run([rdocx_cli, "text", report_docx], check=True).stdout
    assert "This report records the principal inspection" in out
    assert "Figure 1. Cross-section through pier 2" in out


def test_cli_text_json_schema(rdocx_cli, report_docx):
    data = json.loads(run([rdocx_cli, "text", "--json", report_docx], check=True).stdout)
    assert data["schema"] == 1
    paras = data["paragraphs"]
    first = paras[0]
    assert {"body_index", "path", "style", "numbering", "runs", "text"} <= set(first)
    assert {"index", "text", "formatting"} <= set(first["runs"][0])
    # table cell paragraphs are listed too, with a path through the table
    assert any(len(p["path"]) > 1 for p in paras)
    headings = [p["text"] for p in paras if p["style"] == "Heading1"]
    assert headings[:2] == ["Summary", "Scope and method"]


def test_cli_text_includes_block_content_control(rdocx_cli, report_docx):
    out = run([rdocx_cli, "text", report_docx], check=True).stdout
    assert "Access to the river was agreed with the harbour office" in out


def test_cli_inspect_json(rdocx_cli, report_docx):
    data = json.loads(run([rdocx_cli, "inspect", "--json", report_docx], check=True).stdout)
    assert data["tables"] == 3
    assert data["metadata"]["title"].startswith("Riverton Footbridge")
    assert "Heading1" in data["styles_used"]


def test_cli_layout_json(rdocx_cli, report_docx):
    data = json.loads(run([rdocx_cli, "layout", "--json", report_docx], check=True).stdout)
    items = data["body_items"]
    assert items[0]["fragments"][0]["physical_page"] == 1
    pages = {f["physical_page"] for it in items for f in it["fragments"]}
    assert max(pages) >= 10


def test_python_structure(report_docx):
    doc = rdocx.Document.open(report_docx)
    assert len(doc.tables) == 3
    assert len(doc.tables[0].rows) == 57
    assert doc.tables[0].cell(0, 0).text == "Ref"
    sec = doc.sections[0]
    assert sec.orientation == "portrait" and sec.page_width == 7772400 and sec.different_first_page
    ids = {s.style_id for s in doc.styles}
    assert {"Normal", "Heading1", "Caption", "TOC1", "Hyperlink"} <= ids
    kinds = {(v.kind, v.variant) for v in doc.header_footer_variants}
    assert ("footer", "default") in kinds and ("header", "first") in kinds


def test_python_paragraphs_and_runs(report_docx):
    doc = rdocx.Document.open(report_docx)
    heads = [p.text for p in doc.paragraphs if p.style == "Heading1"]
    assert "Inspection findings" in heads
    p = next(p for p in doc.paragraphs if p.text.startswith("The main findings are"))
    assert p.runs[0].text == "The main findings are:"
    bullets = [p for p in doc.paragraphs if p.numbering]
    assert bullets and bullets[0].numbering[0] == 1


def test_python_hyperlinks(report_docx):
    links = rdocx.Document.open(report_docx).hyperlinks
    assert len(links) == 6
    assert all(h.url.startswith("https://example.org/") for h in links)


def test_find_content_indices_includes_toc_block(report_docx):
    """A heading's text is found twice: in the body and in the table of contents block. Never delete by it."""
    doc = rdocx.Document.open(report_docx)
    hits = doc.find_content_indices("Recommendations")
    assert len(hits) == 2 and 7 in hits


def test_python_story_items_cover_headers_and_footers(report_docx):
    doc = rdocx.Document.open(report_docx)
    kinds = {it.story.kind for it in doc.story_items}
    assert {"body", "header", "footer"} <= kinds


def test_python_comments_and_revisions_empty(report_docx):
    doc = rdocx.Document.open(report_docx)
    assert list(doc.comments) == [] and list(doc.revisions) == []


def test_convert_markdown_and_html(rdocx_cli, report_docx, tmp_path):
    md, html = tmp_path / "r.md", tmp_path / "r.html"
    run([rdocx_cli, "convert", report_docx, "--to", "md", "-o", md], check=True)
    run([rdocx_cli, "convert", report_docx, "--to", "html", "-o", html], check=True)
    assert "# Summary" in md.read_text() or "Summary" in md.read_text()
    assert "<table" in html.read_text()


# ---------------------------------------------------------------- stories other than the body, tracked changes
def test_story_items_read_every_story(tmp_path):
    """Headers, footers (each variant), footnotes, text boxes and table cells are read through story_items."""
    path, _ = every_story_docx(tmp_path / "s.docx")
    texts = {(it.story.kind, it.text) for it in rdocx.Document.open(path).story_items if it.text and "NEEDLE" in it.text}
    assert {("table_cell", "Cell NEEDLE"), ("text_box", "Text box NEEDLE"), ("header", "Header NEEDLE"),
            ("footer", "Footer NEEDLE"), ("footer", "First footer NEEDLE"), ("footnote", " Footnote NEEDLE."),
            ("endnote", " Endnote NEEDLE.")} <= texts


def test_text_json_is_the_accepted_view(rdocx_cli, tmp_path):
    """`text --json` and Paragraph.text show tracked insertions and hide deletions."""
    path, _ = every_story_docx(tmp_path / "s.docx")
    data = json.loads(run([rdocx_cli, "text", "--json", path], check=True).stdout)
    assert data["revision_view"] == "accepted"
    assert "Tracked: ins NEEDLE" in [p["text"] for p in data["paragraphs"]]
    assert "Tracked: ins NEEDLE" in [p.text for p in rdocx.Document.open(path).paragraphs]


def test_cli_plain_text_keeps_tracked_insertions(rdocx_cli, tmp_path):
    path, _ = every_story_docx(tmp_path / "s.docx")
    assert "ins NEEDLE" in run([rdocx_cli, "text", path], check=True).stdout


@pytest.mark.parametrize("view", ["plain", "json", "md"])
def test_cli_text_includes_headers_footers_and_notes(rdocx_cli, tmp_path, view):
    path, _ = every_story_docx(tmp_path / "s.docx")
    if view == "plain":
        out = run([rdocx_cli, "text", path], check=True).stdout
        assert "--- header (/word/header1.xml) ---\nHeader NEEDLE\n" in out
    elif view == "json":
        out = run([rdocx_cli, "text", "--json", path], check=True).stdout
        data = json.loads(out)
        assert data["scope"] == "all-supported-stories"
        assert {"header", "footer", "footnote", "endnote", "text_box"} <= {s["kind"] for s in data["stories"]}
    else:
        run([rdocx_cli, "convert", path, "--to", "md", "-o", tmp_path / "s.md"], check=True)
        out = (tmp_path / "s.md").read_text()
    assert "Header NEEDLE" in out and "Footnote NEEDLE" in out and "Text box NEEDLE" in out


# ---------------------------------------------------------------- validation
def broken(src, dst, name, change):
    with zipfile.ZipFile(src) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(dst, "w") as z:
        for info, data in items:
            z.writestr(info, change(data) if info.filename == name else data)
    return dst


def small_with_header(path):
    d = docx.Document()
    d.add_paragraph("x")
    d.sections[0].header.paragraphs[0].text = "Head"
    d.save(path)
    return path


def test_validate_catches_a_truncated_document_part(rdocx_cli, tmp_path):
    src = small_with_header(tmp_path / "a.docx")
    bad = broken(src, tmp_path / "b.docx", "word/document.xml", lambda b: b[: len(b) // 2])
    assert run([rdocx_cli, "validate", bad]).returncode != 0


def test_validate_catches_a_truncated_header(rdocx_cli, tmp_path):
    src = small_with_header(tmp_path / "a.docx")
    with zipfile.ZipFile(src) as z:
        header = next(n for n in z.namelist() if n.startswith("word/header"))
    bad = broken(src, tmp_path / "b.docx", header, lambda b: b[: len(b) // 2])
    assert run([rdocx_cli, "validate", bad]).returncode != 0


def test_validate_catches_a_dangling_style_id(rdocx_cli, tmp_path):
    src = small_with_header(tmp_path / "a.docx")
    bad = broken(src, tmp_path / "b.docx", "word/document.xml",
                 lambda b: b.replace(b"<w:p>", b'<w:p><w:pPr><w:pStyle w:val="NoSuchStyle"/></w:pPr>', 1))
    assert b"NoSuchStyle" in zipfile.ZipFile(bad).read("word/document.xml")
    assert run([rdocx_cli, "validate", bad]).returncode != 0


def test_cli_text_keeps_the_body_when_a_header_is_truncated(rdocx_cli, tmp_path):
    src = small_with_header(tmp_path / "a.docx")
    with zipfile.ZipFile(src) as z:
        header = next(n for n in z.namelist() if n.startswith("word/header"))
    bad = broken(src, tmp_path / "b.docx", header, lambda b: b[: len(b) // 2])
    res = run([rdocx_cli, "text", bad])
    assert res.returncode == 0 and res.stdout.startswith("x") and "Head" not in res.stdout and res.stderr
    data = json.loads(run([rdocx_cli, "text", "--json", bad], check=True).stdout)
    assert (data["scope"], data["stories"]) == ("main", [])


def test_reopen_with_story_items_catches_a_truncated_header(tmp_path):
    """A second check beside `validate`: a story read fails on a truncated header."""
    src = small_with_header(tmp_path / "a.docx")
    with zipfile.ZipFile(src) as z:
        header = next(n for n in z.namelist() if n.startswith("word/header"))
    bad = broken(src, tmp_path / "b.docx", header, lambda b: b[: len(b) // 2])
    with pytest.raises(rdocx.RdocxError):
        rdocx.Document.open(bad).story_items
