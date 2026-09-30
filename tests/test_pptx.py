"""pptx: reading, creating, editing, comments, notes, rendering, layout checks, package round trip."""
import io
import json
import os
import zipfile

import pptx
import pptx.enum.text
import pptx.util
import pytest
import rpptx
from PIL import Image
from rpptx.dml.color import RGBColor
from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE

from conftest import STAMP, digest, parts, run

EMU = 914400


def walk(shapes):
    for sh in shapes:
        yield sh
        if int(sh.shape_type or 0) == 6:  # group
            yield from walk(sh.shapes)


def png_bytes(size=(300, 150), color=(200, 60, 60)):
    buf = io.BytesIO()
    Image.new("RGB", size, color).save(buf, "PNG")
    return buf.getvalue()


# ---------------------------------------------------------------- reading
def test_cli_text_json_and_outline(rpptx_cli, deck_pptx):
    data = json.loads(run([rpptx_cli, "text", "--json", deck_pptx], check=True).stdout)
    s = data["slides"][1]
    assert s["paragraphs"][0]["text"] == "Condition at a glance"
    assert s["paragraphs"][0]["runs"][0]["formatting"]["size_points"] == 30.0
    assert "bearing movement" in s["notes"]
    outline = run([rpptx_cli, "outline", deck_pptx], check=True).stdout
    assert "Slide 5: Programme for option B" in outline


def test_cli_inspect_json(rpptx_cli, deck_pptx):
    data = json.loads(run([rpptx_cli, "inspect", "--json", deck_pptx], check=True).stdout)
    assert data["layouts"] == 11 and len(data["slide_details"]) == 7
    assert "shape_details" in data["slide_details"][1]


def test_new_presentation_is_16_9():
    prs = rpptx.Presentation()
    assert (prs.slide_width, prs.slide_height) == (12192000, 6858000) and len(prs.slide_layouts) == 11


def test_python_traversal_with_groups_and_notes(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    assert (prs.slide_width, prs.slide_height) == (9144000, 6858000)
    texts = [sh.text for sh in walk(prs.slides[1].shapes) if sh.has_text_frame and sh.text]
    assert "74" in texts and "12 mm" in texts  # inside nested groups
    assert prs.slides[3].shapes[1].has_table and prs.slides[3].shapes[1].table.cell(1, 2).text == "212,000"
    assert all(s.notes_text for s in prs.slides)


# ---------------------------------------------------------------- creating
def test_create_and_read_back_with_python_pptx(tmp_path, rpptx_cli):
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.slides[0].shapes.title.text = "Created"
    tb = prs.slides[0].shapes.add_textbox(EMU, 2 * EMU, 4 * EMU, EMU)
    prs.slides[0].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, EMU, 4 * EMU, 2 * EMU, EMU // 2)
    prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, EMU, EMU, 2 * EMU, 2 * EMU)
    prs.slides[0].shapes.add_picture(io.BytesIO(png_bytes()), 5 * EMU, 4 * EMU, width=EMU)
    prs.slides[0].shapes.add_table(2, 2, 5 * EMU, 2 * EMU, 3 * EMU, EMU)
    prs.slides[0].notes_text = "Speaker note."
    box = [s for s in prs.slides[0].shapes if s.has_text_frame and s.name.startswith("TextBox")][0]
    box.text_frame.text = "Body"
    prs.save(tmp_path / "new.pptx")
    x = pptx.Presentation(tmp_path / "new.pptx")
    assert x.slides[0].shapes.title.text == "Created"
    assert x.slides[0].notes_slide.notes_text_frame.text == "Speaker note."
    assert len(x.slides[0].shapes) == 6
    assert run([rpptx_cli, "validate", tmp_path / "new.pptx"]).returncode == 0


# ---------------------------------------------------------------- editing
def test_run_edit_keeps_font(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    r = prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0]
    before = (r.font.name, r.font.size, r.font.color)
    r.text = r.text + " (edited)"
    r = prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0]  # re-fetch: see pptx-run-text-stale
    assert (r.font.name, r.font.size, r.font.color) == before
    assert r.text.endswith("(edited)")


def test_font_color_reads_back_hex(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0].font.color = RGBColor(0x12, 0x34, 0x56)
    assert prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0].font.color == "123456"


def test_units():
    from rpptx.util import Inches, Pt
    assert Inches(1) == EMU and Pt(1) == 12700
    import rpptx.util
    assert not hasattr(rpptx.util, "Emu") and not hasattr(rpptx.util, "Cm")


def test_notes_text_is_none_without_notes():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    assert prs.slides[0].notes_text is None


def test_handles_after_add_and_text_setters():
    """add_* invalidates the handles of every slide; text setters invalidate everything, the new shape
    included; geometry and font setters keep handles valid."""
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides.add_slide(prs.slide_layouts[6])
    a = prs.slides[0].shapes.add_textbox(EMU, EMU, EMU, EMU)
    prs.slides[1].shapes.add_textbox(EMU, EMU, EMU, EMU)
    with pytest.raises(rpptx.StaleElementError):
        a.left
    box = prs.slides[0].shapes.add_textbox(EMU, 2 * EMU, EMU, EMU)
    box.text_frame.text = "x"
    with pytest.raises(rpptx.StaleElementError):
        box.left
    sh = prs.slides[0].shapes[-1]
    sh.left = 2 * EMU
    sh.text_frame.paragraphs[0].runs[0].font.bold = True
    sh.text_frame.paragraphs[0].space_after = 12700
    sh.text_frame.margin_left = 0
    assert sh.left == 2 * EMU
    for setter in (lambda: setattr(prs.slides[0].shapes[-1], "text", "y"),
                   lambda: setattr(prs.slides[0], "notes_text", "note")):
        sh = prs.slides[0].shapes[-1]
        setter()
        with pytest.raises(rpptx.StaleElementError):
            sh.left


def test_run_text_keeps_other_handles_valid(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    shape = prs.slides[1].shapes[1]
    shape.text_frame.paragraphs[0].runs[0].text = "x"
    assert shape.left > 0


def test_paragraph_frame_and_geometry_setters(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    sh = prs.slides[1].shapes[1]
    sh.left, sh.width = sh.left + EMU // 10, sh.width - EMU // 10
    sh = prs.slides[1].shapes[1]
    sh.text_frame.paragraphs[0].space_after = 6 * 12700
    sh = prs.slides[1].shapes[1]
    sh.text_frame.margin_left = 0
    sh = prs.slides[1].shapes[1]
    assert sh.text_frame.paragraphs[0].space_after == 6 * 12700 and sh.text_frame.margin_left == 0


def test_fill_line_and_picture_replace(deck_pptx, tmp_path):
    prs = rpptx.Presentation(deck_pptx)
    sh = prs.slides[1].shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, EMU, EMU, EMU)
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor(0xDD, 0xEE, 0xFF)
    sh.line.width = 12700
    new_png = png_bytes((800, 420), (1, 2, 3))
    pic = next(s for s in prs.slides[0].shapes if int(s.shape_type or 0) == 13)
    pic.replace_image(io.BytesIO(new_png))
    prs.save(tmp_path / "e.pptx")
    x = pptx.Presentation(tmp_path / "e.pptx")
    rect = x.slides[1].shapes[-1]
    assert str(rect.fill.fore_color.rgb) == "DDEEFF" and rect.line.width == 12700
    pic = next(s for s in x.slides[0].shapes if s.shape_type == 13)
    assert pic.image.blob == new_png and pic.image.content_type == "image/png"


def test_slides_notes_hide_move_add_remove(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    n = len(prs.slides)
    prs.slides[2].notes_text = "Edited note."
    prs.slides[3].hidden = True
    prs.slides.move(1, 2)
    prs.slides.add_slide(prs.slide_layouts[1])
    prs.slides.remove(prs.slides[len(prs.slides) - 1])
    assert len(prs.slides) == n and prs.slides[1].notes_text == "Edited note." and prs.slides[3].hidden


def test_cli_replace_expect(rpptx_cli, deck_pptx, tmp_path):
    out = tmp_path / "r.pptx"
    bad = run([rpptx_cli, "replace", deck_pptx, "-p", "option", "-v", "OPTION", "--expect", "1", "-o", out])
    assert bad.returncode != 0 and not out.exists()
    good = run([rpptx_cli, "replace", deck_pptx, "-p", "Riverton Footbridge", "-v", "Kestrel Footbridge", "--expect", "1", "-o", out])
    assert good.returncode == 0, good.stderr
    assert "Kestrel Footbridge" in run([rpptx_cli, "text", out], check=True).stdout


def test_cli_replace_reaches_tables_nested_groups_and_notes(rpptx_cli, deck_pptx, tmp_path):
    for k, (text, n) in enumerate([("212,000", 1), ("12 mm", 2), ("bearing movement", 1)]):
        out = tmp_path / f"r{k}.pptx"
        assert run([rpptx_cli, "replace", deck_pptx, "-p", text, "-v", "X", "--expect", n, "-o", out]).returncode == 0
        assert text not in run([rpptx_cli, "text", "--notes", out], check=True).stdout


def test_duplicate_slide(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    n = len(prs.slides)
    prs.slides.duplicate(prs.slides[1])
    copy = pptx.Presentation(io.BytesIO(prs.to_bytes()))
    assert len(copy.slides) == n + 1
    texts = [[sh.text_frame.text for sh in s.shapes if sh.has_text_frame] for s in copy.slides]
    assert texts.count(texts[1]) == 2


def test_replace_text_python(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    assert prs.try_replace_text("Riverton Footbridge", "Kestrel Footbridge") == 1
    with pytest.raises(rpptx.ReplacementCountError):
        prs.try_replace_text("Kestrel Footbridge", "X", expect=2)


def test_inherited_placeholder_geometry(deck_pptx):
    title = rpptx.Presentation(deck_pptx).slides[1].shapes[0]
    assert title.left is None  # python-pptx reads the inherited value here; rpptx names it effective_geometry
    assert title.effective_geometry() == (457200, 274638, 8229600, 1143000)
    title.left = 914400  # a setter copies the inherited geometry first
    assert (title.left, title.top, title.width, title.height) == (914400, 274638, 8229600, 1143000)


def test_populate_a_new_group():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    g = prs.slides[0].shapes.add_group_shape()
    g.shapes.add_textbox(EMU, EMU, EMU, EMU)
    prs.slides[0].shapes[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, 2 * EMU, 2 * EMU, EMU, EMU)
    group = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[0].shapes[0]
    assert group.shape_type == 6 and len(group.shapes) == 2
    assert (group.left, group.top, group.width, group.height) == (EMU, EMU, 2 * EMU, 2 * EMU)


def test_zorder(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    names = [s.name for s in prs.slides[1].shapes]
    prs.slides[1].shapes.move(3, 0)  # to the back; the last index is the front
    assert [s.name for s in rpptx.Presentation.from_bytes(prs.to_bytes()).slides[1].shapes] == [names[3], *names[:3], *names[4:]]


def test_table_cell_merge_row_height_and_cell_fill(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    table = prs.slides[3].shapes[1].table
    table.cell(0, 0).merge(table.cell(1, 1))
    table.rows[1].height = 600000
    table.cell(2, 0).fill.solid()
    table.cell(2, 0).fill.fore_color.rgb = rpptx.dml.color.RGBColor(0x7B, 0x1E, 0x3A)
    t = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[3].shapes[1].table
    assert t.cell(0, 0).is_merge_origin and (t.cell(0, 0).span_height, t.cell(0, 0).span_width) == (2, 2)
    assert t.rows[1].height == 600000 and str(t.cell(2, 0).fill.fore_color.rgb) == "7B1E3A"


def test_table_add_row(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    rows, cols = len(prs.slides[3].shapes[1].table.rows), len(prs.slides[3].shapes[1].table.columns)
    prs.slides[3].shapes[1].table.rows.add_row()
    prs.slides[3].shapes[1].table.columns.add_column(0)
    prs.slides[3].shapes[1].table.rows.remove(prs.slides[3].shapes[1].table.rows[0])
    t = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[3].shapes[1].table
    assert (len(t.rows), len(t.columns)) == (rows, cols + 1)


# ---------------------------------------------------------------- comments
def test_comments_cli_full_cycle(rpptx_cli, deck_pptx, tmp_path):
    a, b, c, d = (tmp_path / f"{k}.pptx" for k in "abcd")
    run([rpptx_cli, "comment", "add", deck_pptx, "--slide", "2", "--author", "Reviewer", "--text", "Source?",
         "--date", STAMP, "-o", a], check=True)
    thread = json.loads(run([rpptx_cli, "comment", "list", "--json", a], check=True).stdout)["comments"][0]
    assert thread["slide"] == 2 and thread["parent_id"] is None
    run([rpptx_cli, "comment", "reply", a, "--id", thread["id"], "--author", "Author", "--text", "The report.",
         "--date", STAMP, "-o", b], check=True)
    run([rpptx_cli, "comment", "resolve", b, "--id", thread["id"], "-o", c], check=True)
    run([rpptx_cli, "comment", "remove", c, "--id", thread["id"], "-o", d], check=True)
    assert "Source?" not in run([rpptx_cli, "comment", "list", d], check=True).stdout


def test_comments_python_add_and_reply(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    prs.add_comment_author(id="{11111111-2222-3333-4444-555555555555}", name="Reviewer", user_id="reviewer",
                           provider_id="None", initials="R")
    s = prs.slides[1]
    s.add_comment(id="{AAAAAAAA-2222-3333-4444-555555555555}", author_id="{11111111-2222-3333-4444-555555555555}",
                  created=STAMP, text="Source?")
    prs.slides[1].reply_to_comment("{AAAAAAAA-2222-3333-4444-555555555555}", id="{BBBBBBBB-2222-3333-4444-555555555555}",
                                   author_id="{11111111-2222-3333-4444-555555555555}", created=STAMP, text="Report.")
    c = prs.slides[1].comments[0]
    assert c.text == "Source?" and c.replies[0].text == "Report."


def test_comment_resolve_python(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    prs.add_comment_author(id="{11111111-2222-3333-4444-555555555555}", name="Reviewer", user_id="reviewer",
                           provider_id="None", initials="R")
    prs.slides[1].add_comment(id="{AAAAAAAA-2222-3333-4444-555555555555}", author_id="{11111111-2222-3333-4444-555555555555}",
                              created=STAMP, text="Source?")
    prs.slides[1].resolve_comment("{AAAAAAAA-2222-3333-4444-555555555555}")
    assert prs.slides[1].comments[0].status == "resolved"


def test_notes_pdf_and_renders(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    assert prs.to_notes_pdf()[:5] == b"%PDF-" and len(prs.render_all_notes(10)) == len(prs.slides)


def test_slide_handles_hidden_keeps_move_invalidates(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    s = prs.slides[0]
    s.hidden = True
    assert s.hidden is True
    prs.slides.move(0, 1)
    with pytest.raises(rpptx.StaleElementError):
        s.hidden


def test_thumbnail_and_diff(rpptx_cli, deck_pptx, tmp_path):
    run([rpptx_cli, "thumbnail", deck_pptx, "-o", tmp_path / "t.png"], check=True)
    assert Image.open(tmp_path / "t.png").size[0] == 320
    assert run([rpptx_cli, "diff", deck_pptx, deck_pptx], check=True).stdout.strip() == ""


def test_builtin_table_style_is_drawn(tmp_path):
    """python-pptx's default table style is PowerPoint's built-in Medium Style 2 - Accent 1, by GUID: its
    header row is filled with the accent colour (4F81BD in the default theme)."""
    p = pptx.Presentation()
    s = p.slides.add_slide(p.slide_layouts[6])
    s.shapes.add_table(3, 2, EMU, EMU, 4 * EMU, 3 * EMU)
    p.save(tmp_path / "t.pptx")
    png = rpptx.Presentation(tmp_path / "t.pptx").render_slide_to_png(0, 20)
    r, g, b = Image.open(io.BytesIO(png)).convert("RGB").getpixel((int(1.5 * 20), int(1.2 * 20)))
    assert b > r + 40 and b > g + 20


# ---------------------------------------------------------------- layout and rendering
def test_text_layout_finds_the_overflowing_box(deck_pptx):
    frames = rpptx.Presentation(deck_pptx).text_layout()
    assert [(f.slide_index, f.shape_id) for f in frames if f.overflow] == [(4, 19)]


def arial_box(path, paragraphs, width_pt=300, spacing=None):
    """One slide with one text box of 14 pt Arial paragraphs, fixed size, wrapped; `spacing` sets a:spcPct.
    python-pptx's template sets rtl="0" in its text styles."""
    p = pptx.Presentation()
    box = p.slides.add_slide(p.slide_layouts[6]).shapes.add_textbox(EMU, EMU, width_pt * 12700, 100 * 12700)
    box.text_frame.word_wrap, box.text_frame.auto_size = True, pptx.enum.text.MSO_AUTO_SIZE.NONE
    for i, text in enumerate(paragraphs):
        para = box.text_frame.paragraphs[0] if i == 0 else box.text_frame.add_paragraph()
        r = para.add_run()
        r.text, r.font.size, r.font.name = text, pptx.util.Pt(14), "Arial"
        if spacing is not None:
            para.line_spacing = spacing
    p.save(path)
    return path


def test_line_pitch_at_100_percent_covers_the_glyphs(tmp_path):
    """a:spcPct 100 %: one line of 14 pt Arial is 1.2 em (16.8 pt), as in LibreOffice, above the glyphs' 1.117 em."""
    lines = rpptx.Presentation(arial_box(tmp_path / "s.pptx", ["One", "Two"], spacing=1.0)).text_layout()[0].lines
    assert abs(lines[1].baseline - lines[0].baseline - 14 * 1.2) < 0.1


def test_six_lines_at_100_percent_overflow_as_in_libreoffice(tmp_path):
    """Six 14 pt lines at 100 % in a 100 pt box (92.8 pt usable): LibreOffice lays them 100.9 pt high."""
    frame = rpptx.Presentation(arial_box(tmp_path / "s.pptx", ["Line"] * 6, spacing=1.0)).text_layout()[0]
    assert frame.usable.height < 93 and abs(frame.height - 100.9) < 0.5 and frame.overflow


def test_no_line_starts_with_a_comma_a_space_or_a_hyphen(tmp_path):
    """Frame widths from 150 to 350 pt, in a paragraph whose direction is set (rtl="0" in python-pptx's text
    styles): UAX #14 forbids a break before a comma, and LibreOffice breaks after 'in' where a comma follows."""
    text = "Repainted in 3 weeks while still in service, or re-coated next spring"
    prs = rpptx.Presentation(arial_box(tmp_path / "b.pptx", [text]))
    starts = set()
    for width in range(150, 351, 5):
        prs.slides[0].shapes[0].width = width * 12700
        starts |= {line.text[:1] for line in prs.text_layout()[0].lines[1:]}
    assert not starts & {",", " ", "-"}


def test_render_png_and_pdf(deck_pptx, rpptx_cli, tmp_path):
    prs = rpptx.Presentation(deck_pptx)
    png = prs.render_slide_to_png(1, 30)
    assert Image.open(io.BytesIO(png)).size == (300, 225)
    assert prs.to_pdf()[:5] == b"%PDF-"
    run([rpptx_cli, "convert", deck_pptx, "--to", "pdf", "-o", tmp_path / "d.pdf"], check=True)
    assert (tmp_path / "d.pdf").stat().st_size > 10_000


def test_pdf_keeps_the_title_slide_background(deck_pptx, tmp_path):
    import shutil
    import subprocess
    if not shutil.which("pdftoppm"):
        pytest.skip("pdftoppm (poppler) not installed")
    (tmp_path / "d.pdf").write_bytes(rpptx.Presentation(deck_pptx).to_pdf())
    png = subprocess.run(["pdftoppm", "-r", "10", "-png", "-singlefile", "-f", "1", "-l", "1", tmp_path / "d.pdf"],
                         capture_output=True).stdout
    assert Image.open(io.BytesIO(png)).convert("RGB").getpixel((2, 2)) == (123, 30, 58)


def test_open_a_python_pptx_gradient(tmp_path):
    p = pptx.Presentation()
    s = p.slides.add_slide(p.slide_layouts[6])
    s.background.fill.gradient()
    p.save(tmp_path / "g.pptx")
    rpptx.Presentation(tmp_path / "g.pptx")


# ---------------------------------------------------------------- package
def test_noop_save_round_trip(deck_pptx, tmp_path, rpptx_cli):
    rpptx.Presentation(deck_pptx).save(tmp_path / "n.pptx")
    a, b = parts(deck_pptx), parts(tmp_path / "n.pptx")
    assert set(a) == set(b)
    assert run([rpptx_cli, "validate", tmp_path / "n.pptx"]).returncode == 0
    assert len(pptx.Presentation(tmp_path / "n.pptx").slides) == 7


def test_open_a_paragraph_with_two_ppr(tmp_path):
    """a:p children pPr, r, pPr, r: not schema-valid (one pPr, first), met in decks; python-pptx and LibreOffice
    read both runs. rpptx accepts a pPr after a run but refuses the whole file on a second one."""
    p = pptx.Presentation()
    p.slides.add_slide(p.slide_layouts[6]).shapes.add_textbox(EMU, EMU, 4 * EMU, EMU)
    p.save(tmp_path / "a.pptx")
    para = '<a:p><a:pPr algn="l"/><a:r><a:t>One. </a:t></a:r><a:pPr algn="l"/><a:r><a:t>Two.</a:t></a:r></a:p>'
    with zipfile.ZipFile(tmp_path / "a.pptx") as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(tmp_path / "b.pptx", "w") as z:
        for info, data in items:
            if info.filename == "ppt/slides/slide1.xml":
                data = data.replace(b"<a:p/>", para.encode())
            z.writestr(info, data)
    assert pptx.Presentation(tmp_path / "b.pptx").slides[0].shapes[0].text_frame.text == "One. Two."
    assert rpptx.Presentation(tmp_path / "b.pptx").slides[0].shapes[0].text == "One. Two."


def test_convert_refuses_to_overwrite_its_input(rpptx_cli, deck_pptx, copy_of):
    src = copy_of(deck_pptx)
    before = src.read_bytes()
    run([rpptx_cli, "convert", src, "--to", "pdf", "-o", src])
    assert digest(src.read_bytes()) == digest(before)


def test_thumbnail_refuses_to_overwrite_its_input(rpptx_cli, deck_pptx, copy_of):
    src = copy_of(deck_pptx)
    before = src.read_bytes()
    run([rpptx_cli, "thumbnail", src, "-o", src])
    assert digest(src.read_bytes()) == digest(before)


def test_image_convert_and_render_refuse_an_existing_output(rpptx_cli, deck_pptx, tmp_path):
    out = tmp_path / "d.png"
    run([rpptx_cli, "convert", deck_pptx, "--to", "png", "-o", out, "--dpi", "20"], check=True)
    assert (tmp_path / "d_001.png").exists()
    res = run([rpptx_cli, "convert", deck_pptx, "--to", "png", "-o", out, "--dpi", "20"])
    assert res.returncode == 1 and "already exists" in res.stderr
    run([rpptx_cli, "render", deck_pptx, "-o", tmp_path / "s", "--dpi", "20"], check=True)
    res = run([rpptx_cli, "render", deck_pptx, "-o", tmp_path / "s", "--dpi", "20"])
    assert res.returncode == 1 and "already exists" in res.stderr


def test_hyperlink_on_a_run():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_textbox(EMU, EMU, EMU, EMU)
    prs.slides[0].shapes[0].text_frame.text = "link"
    prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address = "https://example.org/"
    run_ = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert run_.hyperlink.address == "https://example.org/"


def test_template_saved_as_presentation_gets_the_presentation_content_type(tmp_path):
    p = pptx.Presentation()
    p.slides.add_slide(p.slide_layouts[6])
    p.save(tmp_path / "a.pptx")
    with zipfile.ZipFile(tmp_path / "a.pptx") as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(tmp_path / "t.potx", "w") as z:
        for info, data in items:
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"presentation.main+xml", b"template.main+xml")
            z.writestr(info, data)
    rpptx.Presentation(tmp_path / "t.potx").save(tmp_path / "b.pptx")
    with zipfile.ZipFile(tmp_path / "b.pptx") as z:
        assert b"presentationml.presentation.main+xml" in z.read("[Content_Types].xml")
