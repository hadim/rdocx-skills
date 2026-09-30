"""docx, fields, table of contents, layout, rendering and CLI robustness."""
import io
import json
import re
import shutil
import subprocess

import docx
import pytest
import rdocx
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt
from PIL import Image

from conftest import STAMP, digest, part, run


def spaced(path, font, size, line, lines=3, picture=None):
    """Paragraphs of one font with explicit spacing (before 0, after 0, `line` auto) in docDefaults."""
    d = docx.Document()
    ppr = d.styles.element.find(qn("w:docDefaults")).find(qn("w:pPrDefault")).find(qn("w:pPr"))
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = ppr.makeelement(qn("w:spacing"), {})
        ppr.append(sp)
    for k, v in (("w:before", "0"), ("w:after", "0"), ("w:line", str(line)), ("w:lineRule", "auto")):
        sp.set(qn(k), v)
    st = d.styles["Normal"]
    st.font.name, st.font.size = font, Pt(size)
    fonts = st.element.get_or_add_rPr().find(qn("w:rFonts"))
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(a), font)
    for i in range(lines):
        d.add_paragraph(f"Line {i} lorem ipsum")
    if picture:
        png = io.BytesIO()
        Image.new("RGB", (400, 400), (200, 200, 200)).save(png, "PNG")
        png.seek(0)
        d.add_picture(png, height=Pt(picture))
        d.add_paragraph("Caption")
    d.save(path)
    return path


# ---------------------------------------------------------------- fields and TOC
def test_page_fields_refresh_fills_cached_results(report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    rep = doc.update_layout_backed_fields()
    assert rep.page_fields >= 1 and rep.num_pages_fields >= 1
    doc.save(tmp_path / "f.docx")
    footer = part(tmp_path / "f.docx", "word/footer1.xml").decode()
    assert re.search(r'fldCharType="separate"/><w:t>\d+</w:t>', footer)


def field_run(p, kind=None, instr=None, text=None):
    r = p.add_run()
    if kind:
        r._r.append(r._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): kind}))
    if instr:
        it = r._r.makeelement(qn("w:instrText"), {})
        it.text = instr
        r._r.append(it)
    if text:
        r.text = text


def test_toc_rebuild_on_a_toc_without_identity_attributes(tmp_path):
    """The cached result of the field spans paragraphs, as Word writes it: begin and separate in the first."""
    d = docx.Document()
    p = d.add_paragraph()
    field_run(p, kind="begin")
    field_run(p, instr=' TOC \\o "1-3" \\h \\z \\u ')
    field_run(p, kind="separate")
    field_run(p, text="Old entry 1")
    q = d.add_paragraph()
    field_run(q, text="Old entry 2")
    field_run(q, kind="end")
    for t in ("Intro", "Method", "Results"):
        d.add_paragraph(t, style="Heading 1")
    d.save(tmp_path / "toc.docx")
    doc = rdocx.Document.open(tmp_path / "toc.docx")
    assert doc.rebuild_toc().entry_count == 3


def test_toc_rebuild_on_a_fresh_open_of_a_word_file(report_docx):
    assert rdocx.Document.open(report_docx).rebuild_toc().entry_count == 21


def test_toc_rebuild_after_an_edit(report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    assert doc.rebuild_toc().entry_count == 21


def test_cli_toc_rebuild(rdocx_cli, tmp_path, report_docx):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    res = run([rdocx_cli, "toc", "rebuild", tmp_path / "e.docx", "-o", tmp_path / "t.docx", "--json"])
    assert res.returncode == 0, res.stderr


def arial_document():
    """python-docx's template (Letter, left and right margins 1.25 in: text from x 90 to x 522 pt), in Arial."""
    d = docx.Document()
    fonts = d.styles["Normal"].element.get_or_add_rPr().get_or_add_rFonts()
    for key in list(fonts.attrib):
        del fonts.attrib[key]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(a), "Arial")
    return d


def pdf_words(pdf, tmp_path):
    """(x_min, x_max, text) of every word of the PDF's first page, in reading order, from its text layer."""
    (tmp_path / "words.pdf").write_bytes(pdf)
    out = subprocess.run(["pdftotext", "-bbox", "-f", "1", "-l", "1", tmp_path / "words.pdf", "-"],
                         capture_output=True, text=True).stdout
    boxes = re.findall(r'xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="[\d.]+">([^<]*)</word>', out)
    return [(float(a), float(b), w) for a, b, w in boxes]


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_toc_entry_of_a_numbered_heading_keeps_its_title_on_the_left(tmp_path):
    """The TOC 1 style carries one right dot-leader tab at the text width and no other stop. The rebuilt entry is
    number, tab, title, tab, page: rebuild_toc gives the entry a left stop after the number, as Word does, and the
    title is laid out there, not at the right stop."""
    d = arial_document()
    toc1 = d.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
    toc1.element.set(qn("w:styleId"), "TOC1")
    toc1.base_style = d.styles["Normal"]
    toc1.element.get_or_add_pPr().append(parse_xml(
        f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>'))
    numbering = d.part.numbering_part.element
    numbering.insert(0, parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="90"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt '
        'w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'))
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
    p = d.add_paragraph()
    field_run(p, kind="begin")
    field_run(p, instr=' TOC \\o "1-1" \\h \\z \\u ')
    field_run(p, kind="separate")
    field_run(p, text="placeholder")
    field_run(d.add_paragraph(), kind="end")
    for title in ("Scope", "Method"):
        h = d.add_paragraph(title, style="Heading 1")
        h._p.get_or_add_pPr().append(parse_xml(
            f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="90"/></w:numPr>'))
    d.save(tmp_path / "toc.docx")
    doc = rdocx.Document.open(tmp_path / "toc.docx")
    assert doc.rebuild_toc().entry_count == 2
    title = next(w for w in pdf_words(doc.to_pdf(), tmp_path) if w[2].startswith("Scope"))
    assert title[0] < 90 + 432 / 2  # left half of the line


def tabbed(path, align, leader="none"):
    """One paragraph "Title<TAB>12" with one custom stop at w:pos 3000 (x 240 pt)."""
    d = arial_document()
    p = d.add_paragraph()
    p._p.get_or_add_pPr().append(parse_xml(
        f'<w:tabs {nsdecls("w")}><w:tab w:val="{align}" w:leader="{leader}" w:pos="3000"/></w:tabs>'))
    for text in ("Title", "\t", "12"):
        p.add_run(text)
    d.save(path)
    return rdocx.Document.open(path).to_pdf()


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_text_after_a_right_tab_ends_at_the_stop(tmp_path):
    """"12" ends at the stop (x 240), as in LibreOffice (240.1)."""
    number = next(w for w in pdf_words(tabbed(tmp_path / "r.docx", "right", "dot"), tmp_path) if w[2].endswith("12"))
    assert abs(number[1] - 240) <= 1


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_text_after_a_left_tab_starts_at_the_stop(tmp_path):
    """"12" starts at the stop (x 240), as in LibreOffice (240.1)."""
    number = next(w for w in pdf_words(tabbed(tmp_path / "l.docx", "left"), tmp_path) if w[2].endswith("12"))
    assert abs(number[0] - 240) <= 1


# ---------------------------------------------------------------- layout
def test_layout_pages_and_fragments(report_docx):
    doc = rdocx.Document.open(report_docx)
    frags = doc.layout()
    assert frags[0].body_index == 0 and frags[0].physical_page == 1
    assert max(f.physical_page for f in frags) == 13  # as LibreOffice
    page = doc.layout_page(0)
    assert page.width == 612.0 and page.height == 792.0


def test_single_line_height_includes_the_line_gap(tmp_path):
    frags = rdocx.Document.open(spaced(tmp_path / "c.docx", "Calibri", 11, 240)).layout()
    assert abs(frags[0].bounds.height - 11 * 1.2207) < 0.2


def test_single_line_height_arial_is_close_to_word(tmp_path):
    frags = rdocx.Document.open(spaced(tmp_path / "a.docx", "Arial", 12, 240)).layout()
    assert abs(frags[0].bounds.height - 12 * 1.150) < 0.2  # Word 1.150 em


def test_picture_line_keeps_the_picture_height(tmp_path):
    frags = rdocx.Document.open(spaced(tmp_path / "p.docx", "Calibri", 11, 264, lines=1, picture=400)).layout()
    assert frags[1].bounds.height < 405  # Word: 402.7 pt from the lead line to the caption


def caption_page(doc):
    paras = {p.text: i for i, p in enumerate(doc.paragraphs)}
    cap = next(t for t in paras if t.startswith("Figure 1."))
    bi = doc.find_content_index(doc.paragraphs[paras[cap]])
    return next(f.physical_page for f in doc.layout() if f.body_index == bi)


def test_figure_caption_stays_with_its_figure_on_the_report(report_docx):
    """Word and LibreOffice keep the caption of figure 1 on page 5."""
    assert caption_page(rdocx.Document.open(report_docx)) == 5


# ---------------------------------------------------------------- rendering
def test_pdf_and_png(report_docx):
    doc = rdocx.Document.open(report_docx)
    pdf = doc.to_pdf()
    assert pdf[:5] == b"%PDF-" and len(pdf) > 100_000
    pngs = doc.render_pages(dpi=36, pages=[0, 1])
    im = Image.open(io.BytesIO(pngs[0]))
    assert len(pngs) == 2 and im.size == (306, 396)


def test_cli_convert_and_render(rdocx_cli, report_docx, tmp_path):
    run([rdocx_cli, "convert", report_docx, "--to", "pdf", "-o", tmp_path / "r.pdf"], check=True)
    assert (tmp_path / "r.pdf").read_bytes()[:5] == b"%PDF-"
    run([rdocx_cli, "render", report_docx, "-o", tmp_path, "--pages", "1-2", "--dpi", "36"], check=True)
    assert len(list(tmp_path.glob("*.png"))) == 2


def redline(path):
    """A word-level redline: "OLDWORD" deleted, "NEWWORD" inserted."""
    a, b = rdocx.Document(), rdocx.Document()
    a.add_paragraph("Keep OLDWORD here.")
    b.add_paragraph("Keep NEWWORD here.")
    a.compare(b, "Reviewer", STAMP, granularity="word")
    a.save(path)
    return a


def pdf_text(pdf, tmp_path):
    (tmp_path / "t.pdf").write_bytes(pdf)
    return subprocess.run(["pdftotext", tmp_path / "t.pdf", "-"], capture_output=True, text=True).stdout


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_pdf_of_a_redline_is_the_accepted_view(rdocx_cli, tmp_path):
    doc = redline(tmp_path / "r.docx")
    assert "Keep NEWWORD here." in pdf_text(doc.to_pdf(), tmp_path) and "OLDWORD" not in pdf_text(doc.to_pdf(), tmp_path)
    run([rdocx_cli, "convert", tmp_path / "r.docx", "--to", "pdf", "-o", tmp_path / "r.pdf"], check=True)
    assert "OLDWORD" not in pdf_text((tmp_path / "r.pdf").read_bytes(), tmp_path)


@pytest.mark.gap("render-tracked-view")
@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_pdf_of_a_redline_can_show_its_revisions(rdocx_cli, tmp_path):
    """The Rust API renders a tracked view (RenderOptions { revision_view: RevisionView::Tracked }: deletions struck
    through, insertions underlined, a change bar); Python and the CLI expose only the accepted view."""
    doc = redline(tmp_path / "r.docx")
    tracked = pdf_text(doc.to_pdf(revision_view="tracked"), tmp_path)
    assert "OLDWORD" in tracked and "NEWWORD" in tracked
    assert len(doc.render_pages(dpi=36, pages=[0], revision_view="tracked")) == 1
    run([rdocx_cli, "convert", tmp_path / "r.docx", "--to", "pdf", "--revision-view", "tracked", "-o",
         tmp_path / "r.pdf"], check=True)
    assert "OLDWORD" in pdf_text((tmp_path / "r.pdf").read_bytes(), tmp_path)


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_pdf_text_layer_arial(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(rdocx.Document.open(spaced(tmp_path / "a.docx", "Arial", 11, 240, lines=1)).to_pdf())
    assert "Line 0 lorem ipsum" in subprocess.run(["pdftotext", pdf, "-"], capture_output=True, text=True).stdout


@pytest.mark.skipif(not shutil.which("pdftotext"), reason="pdftotext (poppler) not installed")
def test_pdf_text_layer_calibri_ligatures(tmp_path):
    d = docx.Document()
    r = d.add_paragraph().add_run("Location Rating Action fifteen office")
    r.font.name = "Calibri"
    d.save(tmp_path / "l.docx")
    (tmp_path / "l.pdf").write_bytes(rdocx.Document.open(tmp_path / "l.docx").to_pdf())
    out = subprocess.run(["pdftotext", tmp_path / "l.pdf", "-"], capture_output=True, text=True).stdout
    assert "Location Rating Action fifteen office" in out


# ---------------------------------------------------------------- CLI robustness and validation
def test_validate_fixture(rdocx_cli, report_docx):
    assert run([rdocx_cli, "validate", report_docx]).returncode == 0


def test_diff_cli(rdocx_cli, report_docx, tmp_path):
    doc = rdocx.Document.open(report_docx)
    doc.try_replace_text("described", "outlined")
    doc.save(tmp_path / "e.docx")
    out = run([rdocx_cli, "diff", report_docx, tmp_path / "e.docx"]).stdout
    assert "outlined" in out


def test_cli_survives_a_closed_pipe(rdocx_cli, report_docx):
    # `text --json` of the report is about 390 kB, far above a pipe buffer, so the write always meets the closed pipe
    p = subprocess.run(f'"{rdocx_cli}" text --json "{report_docx}" 2>/dev/null | head -c 1 > /dev/null; echo ${{PIPESTATUS[0]}}',
                       shell=True, executable="/bin/bash", capture_output=True, text=True)
    assert p.stdout.strip() in ("0", "141")


@pytest.mark.parametrize("fmt", ["pdf", "md", "html"])
def test_convert_refuses_to_overwrite_its_input(rdocx_cli, report_docx, copy_of, fmt):
    src = copy_of(report_docx)
    before = src.read_bytes()
    run([rdocx_cli, "convert", src, "--to", fmt, "-o", src])
    assert digest(src.read_bytes()) == digest(before)


def test_convert_refuses_to_overwrite_the_default_output(rdocx_cli, report_docx, copy_of):
    src = copy_of(report_docx)
    existing = src.with_suffix(".pdf")
    existing.write_bytes(b"keep")
    run([rdocx_cli, "convert", src, "--to", "pdf"])
    assert existing.read_bytes() == b"keep"


def test_convert_writes_the_default_output_next_to_the_input(rdocx_cli, report_docx, copy_of):
    src = copy_of(report_docx)
    run([rdocx_cli, "convert", src, "--to", "pdf"], check=True)
    assert src.with_suffix(".pdf").read_bytes()[:5] == b"%PDF-"


@pytest.mark.parametrize("fmt", ["png", "jpeg", "tiff"])
def test_image_convert_refuses_an_existing_output(rdocx_cli, report_docx, tmp_path, fmt):
    """Several pages go to OUT_001.png, OUT_002.png... (tiff: one file); a second run into them fails."""
    out = tmp_path / f"out.{fmt}"
    run([rdocx_cli, "convert", report_docx, "--to", fmt, "-o", out, "--dpi", "20"], check=True)
    res = run([rdocx_cli, "convert", report_docx, "--to", fmt, "-o", out, "--dpi", "20"])
    assert res.returncode == 1 and "already exists" in res.stderr


def test_image_convert_of_one_page_refuses_its_input(rdocx_cli, tmp_path):
    d = docx.Document()
    d.add_paragraph("One page.")
    d.save(tmp_path / "a.docx")
    before = (tmp_path / "a.docx").read_bytes()
    res = run([rdocx_cli, "convert", tmp_path / "a.docx", "--to", "png", "-o", tmp_path / "a.docx"])
    assert res.returncode == 1 and digest((tmp_path / "a.docx").read_bytes()) == digest(before)


def test_render_refuses_to_write_over_earlier_pages(rdocx_cli, report_docx, tmp_path):
    """Render into a new folder each time: a second render into the same folder fails."""
    out = tmp_path / "pages"
    run([rdocx_cli, "render", report_docx, "-o", out, "--pages", "1", "--dpi", "36"], check=True)
    res = run([rdocx_cli, "render", report_docx, "-o", out, "--pages", "1", "--dpi", "36"])
    assert res.returncode == 1 and "already exists" in res.stderr


def test_editing_subcommands_refuse_an_existing_output(rdocx_cli, report_docx, copy_of):
    src = copy_of(report_docx)
    res = run([rdocx_cli, "replace", src, "-p", "footbridge", "-v", "bridge", "-o", src])
    assert res.returncode != 0 and "already exists" in res.stderr
