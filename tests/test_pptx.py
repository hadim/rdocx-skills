"""pptx: reading, creating, editing, comments, notes, rendering, layout checks, package round trip."""
import io
import re

import pytest
import rpptx
from PIL import Image
from rpptx.dml.color import RGBColor
from rpptx.enum.dml import (MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX)
from rpptx.enum.text import MSO_ANCHOR

from conftest import MONOSPACE_FONTS, digest, font_dir, parts, run

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


# ---------------------------------------------------------------- editing
def test_run_edit_keeps_font(deck_pptx):
    prs = rpptx.Presentation(deck_pptx)
    r = prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0]
    before = (r.font.name, r.font.size, r.font.color.rgb)
    r.text = r.text + " (edited)"
    r = prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0]
    assert (r.font.name, r.font.size, r.font.color.rgb) == before
    assert r.text.endswith("(edited)")


def test_font_color_is_a_colour_format(deck_pptx):
    """As python-pptx: `font.color` is a ColorFormat; `font.color = ...` stays as a shortcut for `.rgb`. Every colour
    setter takes an RGBColor, a hex string with or without '#', or an int triple."""
    prs = rpptx.Presentation(deck_pptx)
    font = prs.slides[1].shapes[1].text_frame.paragraphs[0].runs[0].font
    font.color = RGBColor(0x12, 0x34, 0x56)
    assert type(font.color).__name__ == "ColorFormat" and font.color.rgb == RGBColor(0x12, 0x34, 0x56)
    assert str(font.color.rgb) == "123456" and font.color.type == MSO_COLOR_TYPE.RGB
    for value in ("#654321", "654321", (0x65, 0x43, 0x21)):
        font.color.rgb = value
        assert str(font.color.rgb) == "654321"
    font.color.theme_color = MSO_THEME_COLOR_INDEX.ACCENT_2
    assert (font.color.type, font.color.theme_color) == (MSO_COLOR_TYPE.SCHEME, MSO_THEME_COLOR_INDEX.ACCENT_2)
    shape = prs.slides[1].shapes[1]
    shape.fill.solid()
    shape.fill.fore_color.rgb = "#00FF00"
    shape.line.color.rgb = (0, 0, 255)
    assert (str(shape.fill.fore_color.rgb), str(shape.line.color.rgb)) == ("00FF00", "0000FF")
    with pytest.raises(TypeError):
        font.color.rgb = 0x123456


PPTX_EDITS = {   # edit -> the held handles it retires (slide, shapes, paragraph, run); the others stay valid
    "add_textbox": (lambda p: p.slides[1].shapes.add_textbox(EMU, EMU, EMU, EMU), ""),
    "add_slide": (lambda p: p.slides.add_slide(p.slide_layouts[6]), ""),
    "add_paragraph_and_run": (lambda p: p.slides[0].shapes[0].text_frame.add_paragraph().add_run("x"), ""),
    "geometry_font_paragraph": (lambda p: (setattr(p.slides[0].shapes[0], "left", 2 * EMU),
                                           setattr(p.slides[0].shapes[0].text_frame.paragraphs[0], "space_after", 12700),
                                           setattr(p.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font, "bold", True)), ""),
    "run_text": (lambda p: setattr(p.slides[0].shapes[0].text_frame.paragraphs[0].runs[0], "text", "y"), ""),
    "notes_text": (lambda p: setattr(p.slides[0], "notes_text", "n"), ""),
    "text_frame_text": (lambda p: setattr(p.slides[0].shapes[0].text_frame, "text", "new"), "paragraph run"),
    "shape_text": (lambda p: setattr(p.slides[1].shapes[0], "text", "new"), "paragraph run"),
    "try_replace_text": (lambda p: p.try_replace_text("Draft", "Final", expect=4), "run"),
    "shapes_move": (lambda p: p.slides[1].shapes.move(0, 1), "shapes paragraph run"),
    "shapes_remove": (lambda p: p.slides[1].shapes.remove(p.slides[1].shapes[1]), "shapes paragraph run"),
    "shapes_group": (lambda p: p.slides[1].shapes.group([p.slides[1].shapes[0], p.slides[1].shapes[1]]), "shapes paragraph run"),
    "slides_move": (lambda p: p.slides.move(0, 1), "slide shapes paragraph run"),
    "slides_remove": (lambda p: p.slides.remove(p.slides[1]), "slide shapes paragraph run"),
    "slide_hidden": (lambda p: setattr(p.slides[1], "hidden", True), ""),
}


@pytest.mark.parametrize("op", sorted(PPTX_EDITS))
def test_handles_after_each_kind_of_edit(op):
    """The rule in python-api.md, "Handles": an edit retires the handles of the kind it renumbers and those below
    it (slide > shape > paragraph > run), in every slide; appends and setters retire none."""
    prs = rpptx.Presentation()
    for _ in range(2):
        prs.slides.add_slide(prs.slide_layouts[6])
    for k in range(2):
        prs.slides[k].shapes.add_textbox(EMU, EMU, 4 * EMU, EMU)
        prs.slides[k].shapes.add_textbox(EMU, 3 * EMU, 4 * EMU, EMU)
        prs.slides[k].shapes[0].text_frame.text = "Draft a"
        prs.slides[k].shapes[1].text_frame.text = "Draft b"
    slide, shape = prs.slides[0], prs.slides[0].shapes[1]
    paragraph = prs.slides[0].shapes[1].text_frame.paragraphs[0]
    run_ = paragraph.runs[0]
    held = {"slide": lambda: slide.notes_text, "shapes": lambda: shape.left, "paragraph": lambda: paragraph.text,
            "run": lambda: run_.text}
    edit, retired = PPTX_EDITS[op]
    edit(prs)
    for kind, read in held.items():
        if kind in retired.split():
            with pytest.raises(rpptx.StaleElementError):
                read()
        else:
            read()


def test_a_new_shape_handle_stays_valid_after_its_text_is_set():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    box = prs.slides[0].shapes.add_textbox(EMU, 2 * EMU, EMU, EMU)
    box.text_frame.text = "x"
    box.left = 2 * EMU
    assert prs.slides[0].shapes[0].left == 2 * EMU and box.text == "x"


def blank_slide():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    return prs


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
        sh, r = prs.slides[0].shapes[0], prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        replace()
        sh.left                                        # a replacement retires run handles only
        with pytest.raises(rpptx.StaleElementError):
            r.text
    assert prs.slides[0].shapes[0].text == "H a"


# ---------------------------------------------------------------- gaps reported 2026-10-09
def test_table_cell_text_is_formatted_through_its_text_frame(tmp_path):
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_table(2, 2, rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(4), rpptx.Inches(1.5))
    cell = prs.slides[0].shapes[0].table.cell(0, 0)
    cell.text = "Header"
    font = prs.slides[0].shapes[0].table.cell(0, 0).text_frame.paragraphs[0].runs[0].font
    font.size, font.bold = rpptx.Pt(18), True
    prs.slides[0].shapes[0].table.cell(0, 0).vertical_anchor = MSO_ANCHOR.MIDDLE
    assert prs.slides[0].shapes[0].table.cell(0, 0).vertical_anchor == MSO_ANCHOR.MIDDLE
    prs.save(tmp_path / "t.pptx")
    slide_xml = parts(tmp_path / "t.pptx")["ppt/slides/slide1.xml"].decode()
    assert '<a:tcPr anchor="ctr"' in slide_xml
    assert re.search(r'<a:tc>.*?<a:rPr[^>]*\bsz="1800"[^>]*\bb="1"|<a:tc>.*?<a:rPr[^>]*\bb="1"[^>]*\bsz="1800"', slide_xml, re.S)


def test_run_font_color_rgb_as_in_python_pptx(tmp_path):
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(3), rpptx.Inches(1))
    prs.slides[0].shapes[0].text_frame.paragraphs[0].add_run().text = "red"
    run_ = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    run_.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    assert str(prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color.rgb) == "FF0000"


def test_font_dir_reaches_rpptx_output(rpptx_cli, tmp_path):
    fonts = font_dir(tmp_path, MONOSPACE_FONTS, "Monofamily")
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(rpptx.Inches(1), rpptx.Inches(1), rpptx.Inches(6), rpptx.Inches(1))
    prs.slides[0].shapes[0].text_frame.paragraphs[0].add_run().text = "Sample text in a font from a folder"
    prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.name = "Monofamily"
    prs.save(tmp_path / "f.pptx")
    for args in (["convert", "--to", "pdf"], ["convert", "--to", "png"], ["render"]):
        res = run([rpptx_cli, *args, tmp_path / "f.pptx", "--font-dir", fonts, "-o", tmp_path / args[-1]])
        assert res.returncode == 0, res.stderr
    deck = rpptx.Presentation.from_bytes((tmp_path / "f.pptx").read_bytes())
    assert digest(deck.render_slide_to_png(0, 100.0, font_dir=fonts)) != digest(deck.render_slide_to_png(0, 100.0))
    assert deck.to_pdf(font_dir=fonts) != deck.to_pdf()
    width = [deck.text_layout(**kw)[0].lines[0].bounds.width for kw in ({}, {"font_dir": fonts})]
    assert width[0] != width[1]                        # the fit check measures with the fonts the render draws
