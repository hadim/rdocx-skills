"""pptx: reading, creating, editing, comments, notes, rendering, layout checks, package round trip."""
import io
import json
import os
import re
import zipfile

import pptx
import pptx.enum.text
import pptx.util
import pytest
import rpptx
from PIL import Image
from rpptx.dml.color import RGBColor
from rpptx.enum.dml import MSO_ARROWHEAD_LENGTH, MSO_ARROWHEAD_STYLE, MSO_ARROWHEAD_WIDTH, MSO_LINE_DASH_STYLE
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


def blank_slide():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    return prs


def dark_pixels(png):
    return sum(1 for v in Image.open(io.BytesIO(png)).convert("L").getdata() if v < 200)


def test_add_shape_writes_the_python_pptx_theme_style_and_draws():
    """add_shape writes python-pptx's p:style (accent1 fill and line, effect 2, minor font in lt1), on a
    slide and in a group; add_textbox writes none. Direct fill and line stay unset."""
    ref = pptx.Presentation()
    ref_shapes = ref.slides.add_slide(ref.slide_layouts[6]).shapes
    ref_shapes.add_shape(1, EMU, EMU, 3 * EMU, 2 * EMU)
    ref_shapes.add_group_shape().shapes.add_shape(1, EMU, EMU, EMU, EMU)
    ref_shapes.add_textbox(0, 0, EMU, EMU)
    prs = blank_slide()
    prs.slides[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, EMU, 3 * EMU, 2 * EMU)
    prs.slides[0].shapes.add_group_shape().shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, EMU, EMU, EMU)
    prs.slides[0].shapes.add_textbox(0, 0, EMU, EMU)

    def styles(blob):
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            return re.findall(r"<p:style>.*?</p:style>", z.read("ppt/slides/slide1.xml").decode())

    out = io.BytesIO()
    ref.save(out)
    assert len(styles(prs.to_bytes())) == 2 and styles(prs.to_bytes()) == styles(out.getvalue())
    sh = prs.slides[0].shapes[0]
    assert sh.fill.type is None and sh.line.color.rgb is None and sh.theme_effect_index == 2
    assert dark_pixels(prs.render_slide_to_png(0, 40)) > 0
    sh.theme_effect_index = 0
    assert sh.theme_effect_index == 0 and b'<a:effectRef idx="0">' in sh.xml
    assert prs.slides[0].shapes[2].theme_effect_index is None


def test_connector_theme_effect_line_ends_and_dash():
    prs = blank_slide()
    c = prs.slides[0].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, EMU, EMU, 3 * EMU, EMU)
    assert b'<a:effectRef idx="1">' in c.xml and c.theme_effect_index == 1
    with zipfile.ZipFile(io.BytesIO(prs.to_bytes())) as z:  # effect style 1 of the default theme is a shadow
        theme = z.read("ppt/theme/theme1.xml").decode()
    assert "outerShdw" in theme.split("<a:effectStyleLst>")[1].split("</a:effectStyle>")[0]
    c.theme_effect_index = 0
    c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    c.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
    c.line.tail_end.width = MSO_ARROWHEAD_WIDTH.WIDE
    c.line.tail_end.length = MSO_ARROWHEAD_LENGTH.LONG
    c.line.head_end.type = MSO_ARROWHEAD_STYLE.OVAL
    c.shadow.inherit = False
    assert c.left == EMU  # these setters keep the handle valid
    c = prs.slides[0].shapes[0]
    assert (c.theme_effect_index, c.line.dash_style, c.line.tail_end.type, c.line.head_end.type) == (0, 4, 2, 6)
    xml = c.xml
    assert b'<a:effectRef idx="0">' in xml and b"<a:effectLst/>" in xml
    assert b'<a:tailEnd type="triangle" w="lg" len="lg"/>' in xml and b'<a:headEnd type="oval"/>' in xml
    line = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[0].shapes[0].line
    assert line.dash_style == pptx.enum.dml.MSO_LINE_DASH_STYLE.DASH


def test_shadow_writes_an_outer_shadow():
    prs = blank_slide()
    sh = prs.slides[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, EMU, EMU, EMU)
    assert sh.shadow.inherit and not sh.shadow.visible
    s = sh.shadow
    s.visible = True
    s.color.rgb = RGBColor(0, 0, 0)
    s.alpha = 0.35
    s.blur_radius, s.distance, s.direction, s.align, s.rotate_with_shape = 50800, 38100, 45.0, "tl", False
    assert sh.left == EMU
    s = prs.slides[0].shapes[0].shadow
    assert (s.inherit, s.visible, s.alpha, s.blur_radius, s.distance, s.direction, s.align, s.rotate_with_shape) == (
        False, True, 0.35, 50800, 38100, 45.0, "tl", False)
    x = pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[0].shapes[0]
    outer = x._element.spPr.find(pptx.oxml.ns.qn("a:effectLst")).find(pptx.oxml.ns.qn("a:outerShdw"))
    assert dict(outer.attrib) == {"blurRad": "50800", "dist": "38100", "dir": "2700000", "algn": "tl", "rotWithShape": "0"}
    assert x.shadow.inherit is False


def test_auto_shape_type_changes_the_preset():
    prs = blank_slide()
    prs.slides[0].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, EMU, EMU, EMU, EMU)
    prs.slides[0].shapes.add_textbox(EMU, 3 * EMU, EMU, EMU)
    sh = prs.slides[0].shapes[0]
    assert sh.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
    sh.auto_shape_type = MSO_SHAPE.RECTANGLE
    assert sh.left == EMU
    assert pptx.Presentation(io.BytesIO(prs.to_bytes())).slides[0].shapes[0].auto_shape_type == pptx.enum.shapes.MSO_SHAPE.RECTANGLE
    sh.auto_shape_type = "roundRect"
    assert prs.slides[0].shapes[0].auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE
    with pytest.raises(ValueError):
        prs.slides[0].shapes[1].auto_shape_type


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


def test_import_slide_needs_layout_when_names_differ():
    src = pptx.Presentation()
    src.slides.add_slide(src.slide_layouts[5]).shapes.title.text = "Imported"
    src.slide_layouts[5]._element.cSld.set("name", "Heading only")  # a name the destination lacks
    buf = io.BytesIO()
    src.save(buf)
    source = rpptx.Presentation.from_bytes(buf.getvalue())
    dst = rpptx.Presentation()
    dst.slides.add_slide(dst.slide_layouts[0])
    with pytest.raises(rpptx.RpptxError, match='no layout named "Heading only", pass the destination layout'):
        dst.slides.import_slide(source.slides[0])
    assert len(dst.slides) == 1
    dst.slides.import_slide(source.slides[0], layout=dst.slide_layouts[5], index=0)
    back = pptx.Presentation(io.BytesIO(dst.to_bytes()))
    assert len(back.slides) == 2 and back.slides[0].shapes.title.text == "Imported"
    assert back.slides[0].slide_layout.name == "Title Only"
    same = rpptx.Presentation()  # same layout names: no layout= needed
    same.slides.import_slide(rpptx.Presentation.from_bytes(dst.to_bytes()).slides[0])
    assert same.slides[0].slide_layout.name == "Title Only"


def test_replace_text_in_one_slide_or_one_frame():
    prs = blank_slide()
    prs.slides.add_slide(prs.slide_layouts[6])
    for k in range(2):
        prs.slides[k].shapes.add_textbox(EMU, EMU, 4 * EMU, EMU)
        prs.slides[k].shapes[0].text_frame.text = "Draft a"
        prs.slides[k].notes_text = "Draft note"
    assert prs.slides[1].try_replace_text("Draft", "Final", expect=2) == 2  # the shape and the notes
    assert [s.shapes[0].text for s in prs.slides] == ["Draft a", "Final a"]
    assert [s.notes_text for s in prs.slides] == ["Draft note", "Final note"]
    assert prs.slides[0].try_replace_text("Draft", "Final", expect=1, notes=False) == 1
    assert prs.slides[0].notes_text == "Draft note"
    with pytest.raises(rpptx.ReplacementCountError):
        prs.slides[0].try_replace_text("Final", "X", expect=2)
    with pytest.raises(rpptx.ReplacementCountError):
        prs.slides[0].shapes[0].text_frame.try_replace_text("Final", "X", expect=2)
    assert prs.slides[0].shapes[0].text == "Final a"  # all or nothing
    for replace in (lambda: prs.slides[0].shapes[0].text_frame.try_replace_text("Final", "F", expect=1),
                    lambda: prs.slides[0].try_replace_text("F", "G", expect=1),
                    lambda: prs.try_replace_text("G", "H", expect=1)):
        sh = prs.slides[0].shapes[0]
        replace()
        with pytest.raises(rpptx.StaleElementError):
            sh.left
    assert prs.slides[0].shapes[0].text == "H a"


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
    assert [n for n in a if a[n] != b[n]] == []
    assert run([rpptx_cli, "validate", tmp_path / "n.pptx"]).returncode == 0
    assert len(pptx.Presentation(tmp_path / "n.pptx").slides) == 7


def test_edit_rewrites_only_the_part_it_touches(deck_pptx, tmp_path):
    """A run edit on slide 2 rewrites that slide's part; the other slides, layouts, masters, notes and media keep
    their bytes, and the package gains no part."""
    prs = rpptx.Presentation(deck_pptx)
    run_ = next(p.runs[0] for sh in prs.slides[1].shapes if sh.has_text_frame
                for p in sh.text_frame.paragraphs if p.runs and p.runs[0].text == "Condition at a glance")
    run_.text = "Condition today"
    prs.save(tmp_path / "e.pptx")
    a, b = parts(deck_pptx), parts(tmp_path / "e.pptx")
    assert sorted(b) == sorted(a)
    assert [n for n in a if a[n] != b[n]] == ["ppt/slides/slide2.xml"]


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



def test_shape_click_action_link_and_slide_jump(tmp_path, rpptx_cli):
    """A web link and a slide jump on whole shapes (a group member included); the setters keep handles valid,
    None clears, a jump survives a move, and removing its target leaves a link that does nothing."""
    prs = rpptx.Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides[0].shapes.add_textbox(EMU, EMU, EMU, EMU)
    prs.slides[0].shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, 2 * EMU, EMU, EMU)
    prs.slides[0].shapes.add_group_shape()
    prs.slides[0].shapes[2].shapes.add_shape(MSO_SHAPE.RECTANGLE, EMU, 3 * EMU, EMU, EMU)
    box, rect, slide = prs.slides[0].shapes[0], prs.slides[0].shapes[1], prs.slides[0]
    assert box.click_action.hyperlink.address is None and box.click_action.target_slide is None
    box.click_action.hyperlink.address = "https://example.org/a"
    rect.click_action.target_slide = prs.slides[2]
    assert box.left == EMU and rect.top == 2 * EMU and len(slide.shapes) == 3      # handles still valid
    prs.slides[0].shapes[2].shapes[0].click_action.target_slide = prs.slides[1]
    assert prs.slides[0].shapes[1].click_action.target_slide == prs.slides[2]
    assert prs.slides[0].shapes[1].click_action.hyperlink.address == "slide3.xml"
    assert prs.slides[0].shapes[2].shapes[0].click_action.target_slide == prs.slides[1]
    prs.save(tmp_path / "links.pptx")
    assert run([rpptx_cli, "validate", tmp_path / "links.pptx"]).returncode == 0
    other = pptx.Presentation(tmp_path / "links.pptx")
    shapes = other.slides[0].shapes
    assert shapes[0].click_action.hyperlink.address == "https://example.org/a"
    assert shapes[1].click_action.target_slide.slide_id == other.slides[2].slide_id
    assert shapes[2].shapes[0].click_action.target_slide.slide_id == other.slides[1].slide_id
    prs = rpptx.Presentation(tmp_path / "links.pptx")
    prs.slides[0].shapes[0].click_action.hyperlink.address = None
    assert prs.slides[0].shapes[0].click_action.hyperlink.address is None
    target = prs.slides[2]
    prs.slides.move(2, 1)
    assert prs.slides[0].shapes[1].click_action.target_slide == prs.slides[1]
    prs.slides.remove(prs.slides[1])
    assert prs.slides[0].shapes[1].click_action.target_slide is None
    with pytest.raises(rpptx.StaleElementError):
        target.hidden
    assert b'action="ppaction://noaction"' in parts(io.BytesIO(prs.to_bytes()))["ppt/slides/slide1.xml"]


def test_slide_handles_compare_equal_and_are_unhashable():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.slides.add_slide(prs.slide_layouts[6])
    assert prs.slides[0] == prs.slides[0] and prs.slides[0] != prs.slides[1]
    with pytest.raises(TypeError):
        hash(prs.slides[0])


def test_replace_text_is_try_replace_text_and_validate_is_empty(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    assert prs.validate() == ()
    before = digest(prs.to_bytes())
    with pytest.raises(rpptx.ReplacementCountError):
        prs.replace_text("EUR 230,000", "EUR 236,000", expect=1)
    assert digest(prs.to_bytes()) == before
    assert prs.replace_text("EUR 230,000", "EUR 236,000", expect=2) == 2


def test_comment_anchored_on_a_shape(deck_pptx, rpptx_cli, tmp_path):
    """shape_id anchors a comment on that shape (ac:spMk in the comment part); an unknown shape or author raises
    RpptxError and changes nothing."""
    prs = rpptx.Presentation(deck_pptx)
    author = "{11111111-2222-3333-4444-555555555555}"
    sid = prs.slides[3].shapes[2].shape_id
    kw = dict(id="{AAAAAAAA-2222-3333-4444-555555555555}", author_id=author, created=STAMP)
    before = digest(prs.to_bytes())
    with pytest.raises(rpptx.RpptxError, match="author"):
        prs.slides[3].add_comment(text="x", shape_id=sid, **kw)
    assert digest(prs.to_bytes()) == before
    prs.add_comment_author(id=author, name="Claude", user_id="Claude", provider_id="None", initials="C")
    before = digest(prs.to_bytes())
    with pytest.raises(rpptx.RpptxError, match="shape id 9999"):
        prs.slides[3].add_comment(text="x", shape_id=9999, **kw)
    assert digest(prs.to_bytes()) == before
    prs.slides[3].add_comment(text="On this box?", shape_id=sid, **kw)
    prs.save(tmp_path / "c.pptx")
    assert run([rpptx_cli, "validate", tmp_path / "c.pptx"]).returncode == 0
    comment = next(v for k, v in parts(tmp_path / "c.pptx").items() if k.startswith("ppt/comments/"))
    assert b"<ac:deMkLst" in comment and f'<ac:spMk id="{sid}"/>'.encode() in comment
    listed = json.loads(run([rpptx_cli, "comment", "list", "--json", tmp_path / "c.pptx"], check=True).stdout)
    assert [(c["slide"], c["text"]) for c in listed["comments"]] == [(4, "On this box?")]

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


# ---------------------------------------------------------------- the rest of the API and CLI the references document
def test_lengths_colours_and_record_types(deck_pptx):
    assert isinstance(rpptx.util.Inches(1), rpptx.util.Length) and rpptx.util.Inches(1) == EMU
    assert (rpptx.util.Inches(1).pt, rpptx.util.Pt(12).emu, rpptx.util.Inches(2).inches) == (72.0, 152400, 2.0)
    assert str(RGBColor.from_string("7B1E3A")) == "7B1E3A"
    prs = rpptx.Presentation(deck_pptx)
    prs.add_comment_author(id="{11111111-2222-3333-4444-555555555555}", name="Reviewer", user_id="reviewer",
                           provider_id="None", initials="R")
    author = prs.comment_authors[-1]
    assert isinstance(author, rpptx.CommentAuthor) and (author.name, author.initials, author.user_id) == ("Reviewer", "R", "reviewer")
    prs.slides[1].add_comment(id="{AAAAAAAA-2222-3333-4444-555555555555}", author_id=author.id, created=STAMP, text="Q")
    prs.slides[1].reply_to_comment("{AAAAAAAA-2222-3333-4444-555555555555}", id="{BBBBBBBB-2222-3333-4444-555555555555}",
                                   author_id=author.id, created=STAMP, text="A")
    comment = prs.slides[1].comments[0]
    assert isinstance(comment, rpptx.Comment) and isinstance(comment.replies[0], rpptx.CommentReply)
    assert (comment.replies[0].author_id, comment.replies[0].created, comment.replies[0].text) == (author.id, STAMP, "A")
    line = prs.text_layout()[0].lines[0]
    assert isinstance(line, rpptx.TextLineLayout) and isinstance(line.bounds, rpptx.BoundingBox)
    assert line.bounds.y >= 0 and line.bounds.height > 0


def test_frame_margins_picture_crop_and_cell_format(deck_pptx, tmp_path):
    prs = rpptx.Presentation(deck_pptx)
    box = next(sh for sh in prs.slides[0].shapes if sh.has_text_frame)
    for name, value in (("margin_left", 0), ("margin_right", 12700), ("margin_top", 25400), ("margin_bottom", 38100)):
        setattr(box.text_frame, name, value)
    frame = next(sh for sh in prs.slides[0].shapes if sh.has_text_frame).text_frame
    assert (frame.margin_left, frame.margin_right, frame.margin_top, frame.margin_bottom) == (0, 12700, 25400, 38100)
    prs.slides[0].shapes.add_picture(io.BytesIO(png_bytes()), EMU, EMU, width=EMU)
    pic = list(prs.slides[0].shapes)[-1]
    pic.crop_left, pic.crop_right = 0.1, 0.2
    pic = list(prs.slides[0].shapes)[-1]
    assert (round(pic.crop_left, 3), round(pic.crop_right, 3), pic.crop_top, pic.crop_bottom) == (0.1, 0.2, 0.0, 0.0)
    table = next(sh for sh in prs.slides[3].shapes if sh.has_table).table
    cell = table.cell(0, 0)
    assert cell.is_spanned is False
    cell.margin_left = rpptx.util.Pt(4)
    cell.border_top.width = rpptx.util.Pt(2)
    cell.border_top.color.rgb = RGBColor.from_string("FF0000")
    cell = next(sh for sh in prs.slides[3].shapes if sh.has_table).table.cell(0, 0)
    assert (cell.margin_left, cell.border_top.width, str(cell.border_top.color.rgb)) == (rpptx.util.Pt(4), rpptx.util.Pt(2), "FF0000")
    for name in ("margin_right", "margin_top", "margin_bottom", "border_left", "border_right", "border_bottom"):
        getattr(cell, name)
    prs.save(tmp_path / "f.pptx")
    assert pptx.Presentation(tmp_path / "f.pptx").slides[3].shapes  # still opens in python-pptx


def test_cli_operation_records_force_and_image_options(rpptx_cli, deck_pptx, tmp_path):
    def step(*args, out):
        return json.loads(run([rpptx_cli, *args, "-o", tmp_path / out, "--json"], check=True).stdout)
    cid = step("comment", "add", deck_pptx, "--slide", "1", "--author", "R", "--text", "t", "--date", STAMP, out="a.pptx")["comment_id"]
    reply = step("comment", "reply", tmp_path / "a.pptx", "--id", cid, "--author", "A", "--text", "r", "--date", STAMP, out="b.pptx")
    assert (reply["action"], reply["parent_id"], reply["slide"], reply["schema"]) == ("reply", cid, 1, 1)
    assert step("comment", "resolve", tmp_path / "b.pptx", "--id", cid, out="c.pptx")["action"] == "resolve"
    assert step("comment", "remove", tmp_path / "c.pptx", "--id", cid, out="d.pptx")["comment_id"] == cid
    for args in (["thumbnail", deck_pptx, "-o", tmp_path / "t.png"], ["convert", deck_pptx, "--to", "pdf", "-o", tmp_path / "d.pdf"],
                 ["render", deck_pptx, "-o", tmp_path / "r", "--slide", "1", "--dpi", "20"]):
        run([rpptx_cli, *args], check=True)
        assert run([rpptx_cli, *args]).returncode == 1
        assert run([rpptx_cli, *args, "--force"]).returncode == 0
    run([rpptx_cli, "render", deck_pptx, "-o", tmp_path / "tr", "--slide", "1", "--dpi", "20", "--transparent"], check=True)
    assert Image.open(next((tmp_path / "tr").iterdir())).mode == "RGBA"
    sizes = []
    for quality in ("20", "95"):
        out = tmp_path / f"q{quality}"
        run([rpptx_cli, "render", deck_pptx, "-o", out, "--slide", "2", "--dpi", "40", "--format", "jpeg", "--quality", quality],
            check=True)
        sizes.append(next(out.iterdir()).stat().st_size)
    assert sizes[0] < sizes[1]
