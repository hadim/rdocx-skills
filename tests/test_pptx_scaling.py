"""pptx: every public mutating call of rpptx costs about the same in a deck of N slides and in one of 4N, as in
python-pptx, where an edit touches its element and, for a picture or a slide, its own part. No call grows with the
deck on the pin (tensorbee/rdocx#413 and #416 fixed the 159 that grew on rpptx 0.14.3), so each one guards against
a regression. `test_every_mutating_call_is_timed` keeps the list complete as the API grows.

Each slide of the two decks holds 25 text boxes, 12 rectangles, a 4x4 table, a picture of about 1 MB that does not
compress, and speaker notes: both the shapes and the bytes of the package grow fourfold."""
import ast
import io
import itertools
import time
from pathlib import Path

import pytest
import rpptx
from PIL import ImageFont
from rpptx import CategoryChartData, Inches, Pt, RGBColor, XL_CHART_TYPE
from rpptx.enum.dml import (MSO_ARROWHEAD_LENGTH, MSO_ARROWHEAD_STYLE, MSO_ARROWHEAD_WIDTH, MSO_LINE_DASH_STYLE,
                            MSO_PATTERN_TYPE, MSO_THEME_COLOR)
from rpptx.enum.lang import MSO_LANGUAGE_ID
from rpptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
from rpptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN

from builders import edit_cost_pptx, noise_png

N = 5
CALLS = 3  # fresh targets a call is timed on in each deck, the fastest counting; over 50 ms, two are enough
# A cost independent of the deck measures about 1 time, one that grows with it 2 to 4 times. Between the two bounds,
# the result stays: a call of GROWING counts as fixed under FIXED times, any other one as grown over RATIO times, so
# that the noise of a loaded runner does not turn a call near 2 times into an XPASS or a failure.
RATIO, FIXED, FLOOR_MS = 2.5, 1.5, 5
REOPEN = 16  # timed calls between two openings of a deck, so that the probe slides stay small next to N slides
GROWS = "pptx-edit-cost-grows-with-deck"
AUTHOR = "{6F1B3A52-0000-4000-8000-0000000000A0}"
STAMP = "2026-10-01T12:00:00Z"
LICENSE = rpptx.FontEmbeddingLicense(authorized=True, identity="CC0-1.0: Aileron")
FONT = ImageFont.load_default(12).font_bytes


# ---------------------------------------------------------------- targets: each makes one fresh object on the probe
def box(prs, slide, k):
    shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(0.5))
    shape.text_frame.text = "a"
    return shape


def rect(prs, slide, k):
    return slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))


def rounded(prs, slide, k):
    return slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))


def connector(prs, slide, k):
    return slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT, Inches(1), Inches(1), Inches(2), Inches(2))


def picture(prs, slide, k):
    return slide.shapes.add_picture(io.BytesIO(noise_png(8, k)), Inches(6), Inches(1))


def solid(target):
    def make(prs, slide, k):
        shape = target(prs, slide, k)
        shape.fill.solid()
        return shape
    return make


def gradient(prs, slide, k):
    shape = rect(prs, slide, k)
    shape.fill.gradient()
    return shape


def frame(prs, slide, k):
    return box(prs, slide, k).text_frame


def paragraph(prs, slide, k):
    return frame(prs, slide, k).paragraphs[0]


def bulleted(prs, slide, k):
    p = paragraph(prs, slide, k)
    p.bullet = "•"
    return p


def numbered(prs, slide, k):
    p = paragraph(prs, slide, k)
    p.auto_number = "arabicPeriod"
    return p


def run(prs, slide, k):
    return paragraph(prs, slide, k).runs[0]


def table(prs, slide, k):
    return slide.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(3), Inches(1)).table


def cell(prs, slide, k):
    c = table(prs, slide, k).cell(0, 0)
    c.text = "a"
    return c


def chart_data(values=(1, 2)):
    data = CategoryChartData()
    data.categories = ["a", "b"]
    data.add_series("s", values)
    return data


def chart(prs, slide, k):
    return slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3),
                                  chart_data()).chart


def new_slide(prs, slide, k):
    return prs.slides.add_slide(prs.slide_layouts[6])


def with_comments(count, replies=0):
    """A new slide with `count` comments, the first one with `replies` replies; returns (slide, comment ids)."""
    def make(prs, slide, k):
        if AUTHOR not in [author.id for author in prs.comment_authors]:
            prs.add_comment_author(id=AUTHOR, name="Reviewer", user_id="Reviewer", provider_id="None")
        target, ids = new_slide(prs, slide, k), [guid() for _ in range(count)]
        for c, comment in enumerate(ids):
            target.add_comment(id=comment, author_id=AUTHOR, created=STAMP, text=f"comment {c}")
        for r in range(replies):
            target.reply_to_comment(ids[0], id=guid(), author_id=AUTHOR, created=STAMP, text=f"reply {r}")
        return target, ids
    return make


IDS = itertools.count(1)


def guid(*_):
    """A comment or author id no other call of the run used."""
    return "{6F1B3A52-0000-4000-8000-%012d}" % next(IDS)


def movie(prs, slide, k):
    return slide.shapes.add_movie(io.BytesIO(b"\x00\x00\x00\x18ftypmp42" + bytes(64)), Inches(1), Inches(1),
                                  Inches(2), Inches(2), poster_frame_image=io.BytesIO(noise_png(8, k)),
                                  mime_type="video/mp4")


def source_deck(k=0):
    """A small other deck, its first slide holding one text box; its theme's accent 1 differs with k."""
    source = rpptx.Presentation()
    source.slides.add_slide(source.slide_layouts[6]).shapes.add_textbox(0, 0, Inches(2), Inches(1)).text_frame.text = "x"
    source.slide_master.theme.colors["accent1"] = f"12345{k}"
    return source


# ---------------------------------------------------------------- setups
# A setup takes (prs, probe slide) and returns ready(k), which makes the k-th fresh target untimed and returns the
# timed action; the targets are made one at a time, as the timing needs them.
class FromDeck:
    """A setter value read from the deck being timed."""

    def __init__(self, read):
        self.read = read


LAYOUT_5 = FromDeck(lambda prs: prs.slide_layouts[5])
FIRST_SLIDE = FromDeck(lambda prs: prs.slides[0])


def setter(target, attr, value):
    """`target(...).attr = value` (value[k] for a list) on fresh targets."""
    def setup(prs, slide):
        resolved = value.read(prs) if isinstance(value, FromDeck) else value
        def ready(k):
            obj, v = target(prs, slide, k), resolved[k] if isinstance(resolved, list) else resolved
            return lambda: setattr(obj, attr, v)
        return ready
    return setup


def method(target, name, *args, **kwargs):
    """`target(...).name(*args, **kwargs)` on fresh targets; a callable argument is called with (prs, slide, k)."""
    def setup(prs, slide):
        def ready(k):
            obj = target(prs, slide, k)
            values = [a(prs, slide, k) if callable(a) else a for a in args]
            named = {key: v(prs, slide, k) if callable(v) else v for key, v in kwargs.items()}
            return lambda: getattr(obj, name)(*values, **named)
        return ready
    return setup


def on_made(make, act):
    """act(prs, slide, obj) on objects that make(prs, slide, k) creates."""
    def setup(prs, slide):
        def ready(k):
            obj = make(prs, slide, k)
            return lambda: act(prs, slide, obj)
        return ready
    return setup


def each(setup):
    """A setup written out: `setup(prs, slide)` makes every target first and returns call(k)."""
    def eager(prs, slide):
        call = setup(prs, slide)
        return lambda k: lambda: call(k)
    return eager


def path(*names):
    """A target reached from another one by attribute names: path(rect, "fill", "fore_color")."""
    def make(prs, slide, k):
        obj = names[0](prs, slide, k)
        for name in names[1:]:
            obj = getattr(obj, name)
        return obj
    return make


def presentation(*names):
    return path(lambda prs, slide, k: prs, *names)


def slide_of(*names):
    return path(new_slide, *names)


def siblings(count, name, *args):
    """slide.shapes.name([count new text boxes], *args)."""
    return on_made(lambda prs, slide, k: [box(prs, slide, k) for _ in range(count)],
                   lambda prs, slide, shapes: getattr(slide.shapes, name)(shapes, *args))


def shapes_call(name, *args):
    """slide.shapes.name(*args); a callable argument is called with k."""
    return on_made(lambda prs, slide, k: [a(k) if callable(a) else a for a in args],
                   lambda prs, slide, values: getattr(slide.shapes, name)(*values))


def comment_call(count, replies, call):
    """call(slide, comment ids) on new slides that with_comments(count, replies) makes."""
    return on_made(with_comments(count, replies), lambda prs, slide, made: call(*made))


def media_call(call):
    """call(slide, shape id of a new movie, k)."""
    return on_made(lambda prs, slide, k: (movie(prs, slide, k).shape_id, k),
                   lambda prs, slide, made: call(slide, *made))


def fonts_embedded(prs, names, kinds=("regular",)):
    for name in names:
        for kind in kinds:
            prs.embed_font(name, FONT, kind, license=LICENSE)


KINDS = ["regular", "bold", "italic"]
RGB = RGBColor(0x7B, 0x1E, 0x3A)
URL = "https://example.com/a"
TIMED = {
    # Presentation
    "Presentation.embed_true_type_fonts": setter(presentation(), "embed_true_type_fonts", [True, False, True]),
    "Presentation.save_subset_fonts": setter(presentation(), "save_subset_fonts", [True, False, True]),
    "Presentation.slide_number_start": setter(presentation(), "slide_number_start", [2, 3, 4]),
    "Presentation.slide_width": setter(presentation(), "slide_width", [Inches(11), Inches(12), Inches(13)]),
    "Presentation.slide_height": setter(presentation(), "slide_height", [Inches(7), Inches(7.2), Inches(7.4)]),
    "Presentation.add_comment_author": each(lambda prs, slide: lambda k: prs.add_comment_author(
        id=guid(), name=f"Author {k}", user_id=f"author{k}", provider_id="None")),
    "Presentation.embed_font": each(lambda prs, slide: lambda k: prs.embed_font(
        "Aileron", FONT, KINDS[k], license=LICENSE)),
    "Presentation.remove_embedded_font": each(lambda prs, slide: (fonts_embedded(prs, ["Aileron"], KINDS),
                                                                  lambda k: prs.remove_embedded_font("Aileron", KINDS[k]))[1]),
    "Presentation.remove_font": each(lambda prs, slide: (fonts_embedded(prs, [f"Aileron{k}" for k in range(CALLS)]),
                                                         lambda k: prs.remove_font(f"Aileron{k}"))[1]),
    "Presentation.apply_theme": each(lambda prs, slide: (lambda sources: lambda k: prs.apply_theme(sources[k]))(
        [source_deck(k).to_bytes() for k in range(CALLS)])),
    # Layouts
    "SlideLayout.name": setter(lambda prs, slide, k: prs.slide_layouts[k], "name", [f"Layout {k}" for k in range(CALLS)]),
    "SlideLayout.show_master_shapes": setter(lambda prs, slide, k: prs.slide_layouts[k], "show_master_shapes", False),
    "SlideLayout.follow_master_background": setter(lambda prs, slide, k: prs.slide_layouts[k],
                                                   "follow_master_background", False),
    "SlideLayoutCollection.duplicate": each(lambda prs, slide: lambda k: prs.slide_layouts.duplicate(prs.slide_layouts[k])),
    "SlideLayoutCollection.remove": on_made(lambda prs, slide, k: prs.slide_layouts.duplicate(prs.slide_layouts[k]),
                                            lambda prs, slide, layout: prs.slide_layouts.remove(layout)),
    # Slides
    "Slide.show_master_shapes": setter(new_slide, "show_master_shapes", False),
    "Slide.notes_text": setter(new_slide, "notes_text", "notes"),
    "Slide.slide_layout": setter(new_slide, "slide_layout", LAYOUT_5),
    "Slide.hidden": setter(new_slide, "hidden", True),
    "Slide.follow_master_background": setter(new_slide, "follow_master_background", False),
    "Slide.try_replace_text": method(lambda prs, slide, k: (box(prs, slide, k), slide)[1], "try_replace_text", "a", "b"),
    "Slide.add_comment": comment_call(0, 0, lambda s, ids: s.add_comment(id=guid(), author_id=AUTHOR, created=STAMP,
                                                                         text="new")),
    "Slide.reply_to_comment": comment_call(1, 0, lambda s, ids: s.reply_to_comment(
        ids[0], id=guid(), author_id=AUTHOR, created=STAMP, text="reply")),
    "Slide.resolve_comment": comment_call(1, 0, lambda s, ids: s.resolve_comment(ids[0])),
    "Slide.remove_comment": comment_call(1, 0, lambda s, ids: s.remove_comment(ids[0])),
    "Slide.move_comment": comment_call(2, 0, lambda s, ids: s.move_comment(0, 1)),
    "Slide.move_reply": comment_call(1, 2, lambda s, ids: s.move_reply(ids[0], 0, 1)),
    "Slide.replace_media": media_call(lambda slide, shape_id, k: slide.replace_media(
        shape_id, io.BytesIO(b"\x00\x00\x00\x18ftypmp42" + bytes([k]) * 64), mime_type="video/mp4")),
    "Slide.remove_media": media_call(lambda slide, shape_id, k: slide.remove_media(shape_id)),
    "NotesTextFrame.text": setter(lambda prs, slide, k: new_slide(prs, slide, k).notes_slide.notes_text_frame,
                                  "text", "notes"),
    "SlideCollection.add_slide": each(lambda prs, slide: lambda k: prs.slides.add_slide(prs.slide_layouts[6])),
    "SlideCollection.add_slide_at": each(lambda prs, slide: lambda k: prs.slides.add_slide_at(prs.slide_layouts[6], 0)),
    "SlideCollection.import_slide": each(lambda prs, slide: (lambda source: lambda k: prs.slides.import_slide(
        source.slides[0], layout=prs.slide_layouts[6]))(source_deck())),
    "SlideCollection.import_slide_with_diagnostics": each(lambda prs, slide: (
        lambda source: lambda k: prs.slides.import_slide_with_diagnostics(source.slides[0],
                                                                          layout=prs.slide_layouts[6]))(source_deck())),
    "SlideCollection.import_slides": each(lambda prs, slide: (lambda source: lambda k: prs.slides.import_slides(
        source, [0], layout=prs.slide_layouts[6]))(source_deck())),
    "SlideCollection.remove": on_made(new_slide, lambda prs, slide, target: prs.slides.remove(target)),
    "SlideCollection.duplicate": on_made(new_slide, lambda prs, slide, target: prs.slides.duplicate(target)),
    "SlideCollection.move": each(lambda prs, slide: lambda k: prs.slides.move(len(prs.slides) - 1, 0)),
    "SlideTransition.type": setter(slide_of("transition"), "type", "fade"),
    "SlideTransition.direction": setter(lambda prs, slide, k: (lambda t: (setattr(t, "type", "push"), t)[1])(
        new_slide(prs, slide, k).transition), "direction", "up"),
    "SlideTransition.duration": setter(lambda prs, slide, k: (lambda t: (setattr(t, "type", "fade"), t)[1])(
        new_slide(prs, slide, k).transition), "duration", 1.5),
    "SlideTransition.advance_on_click": setter(slide_of("transition"), "advance_on_click", False),
    "SlideTransition.advance_after": setter(slide_of("transition"), "advance_after", 2.0),
    "HeaderFooter.slide_number": setter(slide_of("header_footer"), "slide_number", True),
    "HeaderFooter.footer": setter(slide_of("header_footer"), "footer", "footer"),
    "HeaderFooter.date": setter(slide_of("header_footer"), "date", "1 October 2026"),
    "HeaderFooter.date_format": setter(lambda prs, slide, k: (lambda h: (setattr(h, "date", "auto"), h)[1])(
        new_slide(prs, slide, k).header_footer), "date_format", "datetime2"),
    "FillFormat.solid (slide background)": method(slide_of("background", "fill"), "solid"),
    # Master and theme
    "Theme.name": setter(presentation("slide_master", "theme"), "name", [f"Theme {k}" for k in range(CALLS)]),
    "ThemeColors.__setitem__": each(lambda prs, slide: lambda k: prs.slide_master.theme.colors.__setitem__(
        "accent1", f"11223{k}")),
    "ThemeFontSet.latin": setter(presentation("slide_master", "theme", "fonts", "major"), "latin",
                                 ["Arial", "Georgia", "Verdana"]),
    "ThemeFontSet.east_asian": setter(presentation("slide_master", "theme", "fonts", "major"), "east_asian",
                                      ["MS Gothic", "SimSun", "Meiryo"]),
    "ThemeFontSet.complex_script": setter(presentation("slide_master", "theme", "fonts", "major"), "complex_script",
                                          ["Arial", "Tahoma", "Mangal"]),
    "Font.size (master text style)": setter(lambda prs, slide, k: prs.slide_master.text_styles.title.level(k + 1).font,
                                            "size", Pt(30)),
    # Shapes
    "ShapeCollection.add_textbox": shapes_call("add_textbox", Inches(5), Inches(5), Inches(1), Inches(0.4)),
    "ShapeCollection.add_shape": shapes_call("add_shape", MSO_SHAPE.OVAL, Inches(5), Inches(1), Inches(1), Inches(1)),
    "ShapeCollection.add_connector": shapes_call("add_connector", MSO_CONNECTOR_TYPE.STRAIGHT, 0, 0, Inches(1), Inches(1)),
    "ShapeCollection.add_group_shape": shapes_call("add_group_shape"),
    "ShapeCollection.add_table": shapes_call("add_table", 2, 2, Inches(1), Inches(5), Inches(3), Inches(1)),
    "ShapeCollection.add_picture": shapes_call("add_picture", lambda k: io.BytesIO(noise_png(8, 200 + k)), Inches(6),
                                               Inches(1)),
    "ShapeCollection.add_chart": shapes_call("add_chart", XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1),
                                             Inches(4), Inches(3), lambda k: chart_data()),
    "ShapeCollection.add_movie": each(lambda prs, slide: lambda k: movie(prs, slide, k)),
    "ShapeCollection.group": siblings(2, "group"),
    "ShapeCollection.align": siblings(2, "align", "left"),
    "ShapeCollection.distribute": siblings(3, "distribute", "horizontal"),
    "ShapeCollection.remove": on_made(box, lambda prs, slide, shape: slide.shapes.remove(shape)),
    "ShapeCollection.move": each(lambda prs, slide: (box(prs, slide, 0), box(prs, slide, 1),
                                                     lambda k: slide.shapes.move(len(slide.shapes) - 1, 0))[2]),
    "Shape.alt_text": setter(rect, "alt_text", "x"),
    "Shape.alt_title": setter(rect, "alt_title", "x"),
    "Shape.decorative": setter(rect, "decorative", True),
    "Shape.left": setter(rect, "left", Inches(2)),
    "Shape.top": setter(rect, "top", Inches(2)),
    "Shape.width": setter(rect, "width", Inches(2)),
    "Shape.height": setter(rect, "height", Inches(2)),
    "Shape.flip_h": setter(rect, "flip_h", True),
    "Shape.flip_v": setter(rect, "flip_v", True),
    "Shape.begin_x": setter(connector, "begin_x", Inches(3)),
    "Shape.begin_y": setter(connector, "begin_y", Inches(3)),
    "Shape.end_x": setter(connector, "end_x", Inches(4)),
    "Shape.end_y": setter(connector, "end_y", Inches(4)),
    "Shape.name": setter(rect, "name", [f"Shape {k}" for k in range(CALLS)]),
    "Shape.rotation": setter(rect, "rotation", 45.0),
    "Shape.auto_shape_type": setter(rect, "auto_shape_type", MSO_SHAPE.OVAL),
    "Shape.theme_effect_index": setter(rect, "theme_effect_index", 1),
    "Shape.crop_left": setter(picture, "crop_left", 0.1),
    "Shape.crop_top": setter(picture, "crop_top", 0.1),
    "Shape.crop_right": setter(picture, "crop_right", 0.1),
    "Shape.crop_bottom": setter(picture, "crop_bottom", 0.1),
    "Shape.text": setter(rect, "text", "x"),
    "Shape.begin_connect": method(connector, "begin_connect", rect, 0),
    "Shape.end_connect": method(connector, "end_connect", rect, 2),
    "Shape.ungroup": method(lambda prs, slide, k: slide.shapes.group([box(prs, slide, k), box(prs, slide, k)]), "ungroup"),
    "Shape.remove_style": method(rect, "remove_style"),
    "Shape.insert_picture": method(lambda prs, slide, k: prs.slides.add_slide(prs.slide_layouts[8]).placeholders[1],
                                   "insert_picture", lambda prs, slide, k: io.BytesIO(noise_png(8, 300 + k))),
    "Shape.replace_image": method(picture, "replace_image", lambda prs, slide, k: io.BytesIO(noise_png(8, 400 + k))),
    "AdjustmentCollection.__setitem__": method(path(rounded, "adjustments"), "__setitem__", 0, 0.3),
    "ShapeClickAction.target_slide": setter(path(rect, "click_action"), "target_slide", FIRST_SLIDE),
    "ShapeHyperlink.address": setter(path(rect, "click_action", "hyperlink"), "address", URL),
    "Element.replace_xml (shape text)": each(lambda prs, slide: (lambda shapes: lambda k: shapes[k].element.replace_xml(
        shapes[k].element.xml.replace(">a<", ">b<")))([box(prs, slide, k) for k in range(CALLS)])),
    "Element.replace_xml (slide)": each(lambda prs, slide: (lambda slides: lambda k: slides[k].element.replace_xml(
        slides[k].element.xml.replace("<p:cSld>", '<p:cSld name="x">')))(
            [new_slide(prs, slide, k) for k in range(CALLS)])),
    # Fills, lines, colours, shadows
    "FillFormat.solid": method(path(rect, "fill"), "solid"),
    "FillFormat.background": method(path(rect, "fill"), "background"),
    "FillFormat.gradient": method(path(rect, "fill"), "gradient"),
    "FillFormat.patterned": method(path(rect, "fill"), "patterned"),
    "FillFormat.picture": method(path(rect, "fill"), "picture", lambda prs, slide, k: io.BytesIO(noise_png(8, 500 + k))),
    "FillFormat.transparency": setter(path(solid(rect), "fill"), "transparency", 0.5),
    "FillFormat.pattern": setter(lambda prs, slide, k: (lambda f: (f.patterned(), f)[1])(rect(prs, slide, k).fill),
                                 "pattern", MSO_PATTERN_TYPE.CROSS),
    "FillFormat.gradient_angle": setter(path(gradient, "fill"), "gradient_angle", 45.0),
    "FillFormat.gradient_path": setter(path(gradient, "fill"), "gradient_path", "circle"),
    "FillFormat.solid (table cell)": method(path(cell, "fill"), "solid"),
    "FillFormat.background (table cell)": method(path(cell, "fill"), "background"),
    "FillFormat.solid (line)": method(path(rect, "line", "fill"), "solid"),
    "FillFormat.solid (cell border)": method(path(cell, "border_left", "fill"), "solid"),
    "GradientStops.append": method(path(gradient, "fill", "gradient_stops"), "append", 0.5),
    "GradientStops.__delitem__": each(lambda prs, slide: (lambda stops: lambda k: stops[k].__delitem__(0))(
        [(lambda s: (s.append(0.5), s)[1])(gradient(prs, slide, k).fill.gradient_stops) for k in range(CALLS)])),
    "GradientStop.position": setter(lambda prs, slide, k: gradient(prs, slide, k).fill.gradient_stops[0], "position", 0.3),
    "LineFormat.width": setter(path(rect, "line"), "width", Pt(2)),
    "LineFormat.dash_style": setter(path(rect, "line"), "dash_style", MSO_LINE_DASH_STYLE.DASH),
    "LineFormat.width (cell border)": setter(path(cell, "border_left"), "width", Pt(2)),
    "LineEndFormat.type": setter(path(connector, "line", "head_end"), "type", MSO_ARROWHEAD_STYLE.TRIANGLE),
    "LineEndFormat.width": setter(path(connector, "line", "head_end"), "width", MSO_ARROWHEAD_WIDTH.WIDE),
    "LineEndFormat.length": setter(path(connector, "line", "head_end"), "length", MSO_ARROWHEAD_LENGTH.LONG),
    "ColorFormat.rgb": setter(path(solid(rect), "fill", "fore_color"), "rgb", RGB),
    "ColorFormat.theme_color": setter(path(solid(rect), "fill", "fore_color"), "theme_color", MSO_THEME_COLOR.ACCENT_2),
    "ColorFormat.brightness": setter(lambda prs, slide, k: (lambda c: (setattr(c, "theme_color", MSO_THEME_COLOR.ACCENT_2),
                                                                       c)[1])(solid(rect)(prs, slide, k).fill.fore_color),
                                     "brightness", 0.4),
    "ColorFormat.rgb (font)": setter(path(run, "font", "color"), "rgb", RGB),
    "ShadowFormat.inherit": setter(path(rect, "shadow"), "inherit", False),
    "ShadowFormat.visible": setter(path(rect, "shadow"), "visible", True),
    "ShadowFormat.alpha": setter(path(rect, "shadow"), "alpha", 0.3),
    "ShadowFormat.blur_radius": setter(path(rect, "shadow"), "blur_radius", Pt(4)),
    "ShadowFormat.distance": setter(path(rect, "shadow"), "distance", Pt(3)),
    "ShadowFormat.direction": setter(path(rect, "shadow"), "direction", 45.0),
    "ShadowFormat.align": setter(path(rect, "shadow"), "align", "tl"),
    "ShadowFormat.rotate_with_shape": setter(path(rect, "shadow"), "rotate_with_shape", False),
    # Text
    "TextFrame.text": setter(frame, "text", "x"),
    "TextFrame.clear": method(lambda prs, slide, k: (lambda f: (f.add_paragraph(), f)[1])(frame(prs, slide, k)), "clear"),
    "TextFrame.try_replace_text": method(frame, "try_replace_text", "a", "b"),
    "TextFrame.fit_text": method(frame, "fit_text", fonts=[("Aileron", FONT)], font_family="Aileron"),
    "TextFrame.add_paragraph": method(frame, "add_paragraph"),
    "TextFrame.auto_size": setter(frame, "auto_size", MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT),
    "TextFrame.margin_left": setter(frame, "margin_left", Inches(0.3)),
    "TextFrame.margin_right": setter(frame, "margin_right", Inches(0.3)),
    "TextFrame.margin_top": setter(frame, "margin_top", Inches(0.3)),
    "TextFrame.margin_bottom": setter(frame, "margin_bottom", Inches(0.3)),
    "TextFrame.vertical_anchor": setter(frame, "vertical_anchor", MSO_ANCHOR.MIDDLE),
    "TextFrame.word_wrap": setter(frame, "word_wrap", False),
    "Paragraph.text": setter(paragraph, "text", "x"),
    "Paragraph.level": setter(paragraph, "level", 1),
    "Paragraph.add_field": method(paragraph, "add_field", "slidenum"),
    "Paragraph.add_run": method(paragraph, "add_run", "x"),
    "Paragraph.alignment": setter(paragraph, "alignment", PP_ALIGN.CENTER),
    "Paragraph.line_spacing": setter(paragraph, "line_spacing", 1.5),
    "Paragraph.space_before": setter(paragraph, "space_before", Pt(6)),
    "Paragraph.space_after": setter(paragraph, "space_after", Pt(6)),
    "Paragraph.left_indent": setter(paragraph, "left_indent", Inches(0.5)),
    "Paragraph.right_indent": setter(paragraph, "right_indent", Inches(0.5)),
    "Paragraph.first_line_indent": setter(paragraph, "first_line_indent", Inches(0.2)),
    "Paragraph.bullet": setter(paragraph, "bullet", "•"),
    "Paragraph.bullet_char": setter(paragraph, "bullet_char", "-"),
    "Paragraph.auto_number": setter(paragraph, "auto_number", "arabicPeriod"),
    "Paragraph.auto_number_start": setter(numbered, "auto_number_start", 3),
    "Paragraph.bullet_color": setter(bulleted, "bullet_color", RGB),
    "Paragraph.bullet_size": setter(bulleted, "bullet_size", 1.2),
    "Paragraph.bullet_font": setter(bulleted, "bullet_font", "Arial"),
    "Font.bold (paragraph)": setter(path(paragraph, "font"), "bold", True),
    "Run.text": setter(run, "text", "x"),
    "Hyperlink.address": setter(path(run, "hyperlink"), "address", URL),
    "Hyperlink.target_slide": setter(path(run, "hyperlink"), "target_slide", FIRST_SLIDE),
    "Font.name": setter(path(run, "font"), "name", "Arial"),
    "Font.color": setter(path(run, "font"), "color", RGB),
    "Font.highlight_color": setter(path(run, "font"), "highlight_color", RGB),
    "Font.highlight": setter(path(run, "font"), "highlight", RGB),
    "Font.bold": setter(path(run, "font"), "bold", True),
    "Font.italic": setter(path(run, "font"), "italic", True),
    "Font.underline": setter(path(run, "font"), "underline", True),
    "Font.strike": setter(path(run, "font"), "strike", True),
    "Font.all_caps": setter(path(run, "font"), "all_caps", True),
    "Font.small_caps": setter(path(run, "font"), "small_caps", True),
    "Font.size": setter(path(run, "font"), "size", Pt(20)),
    "Font.baseline": setter(path(run, "font"), "baseline", 0.3),
    "Font.spacing": setter(path(run, "font"), "spacing", Pt(2)),
    "Font.language_id": setter(path(run, "font"), "language_id", MSO_LANGUAGE_ID.FRENCH),
    "Font.language": setter(path(run, "font"), "language", "de-DE"),
    "Font.east_asian_name": setter(path(run, "font"), "east_asian_name", "MS Gothic"),
    "Font.complex_script_name": setter(path(run, "font"), "complex_script_name", "Arial"),
    # Tables
    "Table.first_row": setter(table, "first_row", False),
    "Table.last_row": setter(table, "last_row", True),
    "Table.first_col": setter(table, "first_col", True),
    "Table.last_col": setter(table, "last_col", True),
    "Table.horz_banding": setter(table, "horz_banding", False),
    "Table.vert_banding": setter(table, "vert_banding", True),
    "Table.style_id": setter(table, "style_id", "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"),
    "ColumnCollection.add_column": method(path(table, "columns"), "add_column"),
    "ColumnCollection.remove": each(lambda prs, slide: (lambda tables: lambda k: tables[k].columns.remove(
        tables[k].columns[0]))([table(prs, slide, k) for k in range(CALLS)])),
    "Column.width": setter(lambda prs, slide, k: table(prs, slide, k).columns[0], "width", Inches(2)),
    "RowCollection.add_row": method(path(table, "rows"), "add_row"),
    "RowCollection.remove": each(lambda prs, slide: (lambda tables: lambda k: tables[k].rows.remove(
        tables[k].rows[0]))([table(prs, slide, k) for k in range(CALLS)])),
    "Row.height": setter(lambda prs, slide, k: table(prs, slide, k).rows[0], "height", Inches(1)),
    "Cell.vertical_anchor": setter(cell, "vertical_anchor", MSO_ANCHOR.MIDDLE),
    "Cell.text": setter(cell, "text", "x"),
    "Cell.merge": each(lambda prs, slide: (lambda tables: lambda k: tables[k].cell(0, 0).merge(tables[k].cell(0, 1)))(
        [table(prs, slide, k) for k in range(CALLS)])),
    "Cell.split": each(lambda prs, slide: (lambda tables: lambda k: tables[k].cell(0, 0).split())(
        [(lambda t: (t.cell(0, 0).merge(t.cell(0, 1)), t)[1])(table(prs, slide, k)) for k in range(CALLS)])),
    "Cell.margin_left": setter(cell, "margin_left", Inches(0.2)),
    "Cell.margin_right": setter(cell, "margin_right", Inches(0.2)),
    "Cell.margin_top": setter(cell, "margin_top", Inches(0.2)),
    "Cell.margin_bottom": setter(cell, "margin_bottom", Inches(0.2)),
    # Charts
    "Chart.has_legend": setter(chart, "has_legend", True),
    "Chart.legend_position": setter(lambda prs, slide, k: (lambda c: (setattr(c, "has_legend", True), c)[1])(
        chart(prs, slide, k)), "legend_position", "b"),
    "Chart.legend_overlay": setter(lambda prs, slide, k: (lambda c: (setattr(c, "has_legend", True), c)[1])(
        chart(prs, slide, k)), "legend_overlay", True),
    "Chart.replace_data": method(chart, "replace_data", lambda prs, slide, k: chart_data((3, 4))),
    # Document properties
    "CustomProperties.set": each(lambda prs, slide: lambda k: prs.custom_properties.set(f"key{k}", "value")),
    "CustomProperties.delete": each(lambda prs, slide: ([prs.custom_properties.set(f"gone{k}", "v") for k in range(CALLS)],
                                                        lambda k: prs.custom_properties.delete(f"gone{k}"))[1]),
}
CORE = {"created": [f"2026-10-0{k + 1}T12:00:00Z" for k in range(CALLS)], "revision": [k + 2 for k in range(CALLS)]}
CORE["modified"] = CORE["last_printed"] = CORE["created"]
for name in ("author", "category", "comments", "content_status", "created", "identifier", "keywords", "language",
             "last_modified_by", "last_printed", "modified", "revision", "subject", "title", "version"):
    TIMED[f"CoreProperties.{name}"] = setter(presentation("core_properties"), name,
                                             CORE.get(name, [f"{name} {k}" for k in range(CALLS)]))
for name in ("template", "manager", "company", "presentation_format", "application", "application_version"):
    TIMED[f"AppProperties.{name}"] = setter(presentation("app_properties"), name,
                                            [f"{name} {k}" if name != "application_version" else f"16.{k}000"
                                             for k in range(CALLS)])
for name in ("pages", "words", "characters", "lines", "paragraphs", "slides", "notes", "total_time", "hidden_slides",
             "multimedia_clips", "characters_with_spaces"):
    TIMED[f"AppProperties.{name}"] = setter(presentation("app_properties"), name, [k + 7 for k in range(CALLS)])
for name in ("scale_crop", "links_up_to_date", "shared_document", "hyperlinks_changed"):
    TIMED[f"AppProperties.{name}"] = setter(presentation("app_properties"), name, [True, False, True])

# Calls that work on every slide, or on the whole package, by design: their cost grows with the deck in python-pptx too.
WHOLE_DECK = {
    "Presentation.sign": "signs every part of the package",
    "Presentation.set_header_footer": "applies to every slide",
    "Presentation.set_sections": "lists every slide",
    "Presentation.try_replace_text": "searches every slide",
    "Presentation.replace_text": "searches every slide",
    "Presentation.replace_all": "searches every slide",
    "Presentation.refresh_autofit": "lays out every text frame",
    "SlideTransition.apply_to_all": "applies to every slide",
}
# Public methods that change nothing in the deck.
READS = {"Presentation.export_html", "Presentation.export_markdown", "Presentation.fit_report", "Presentation.save",
         "Presentation.save_html", "Presentation.save_markdown", "Presentation.save_odp", "Presentation.text_layout",
         "Presentation.validate", "Presentation.verify_signatures", "Slide.extract_media", "Shape.effective_geometry",
         "Table.cell", "ThemeColors.keys", "TextStyleList.level", "CustomProperties.get", "SlideLayoutCollection.index",
         "SlideCollection.index"}


def mutating_surface():
    """Class.member for every public method, property setter, __setitem__ and __delitem__ of the rpptx stubs, but the
    outputs (render*, to_*) and READS."""
    tree = ast.parse((Path(rpptx.__file__).parent / "_rpptx.pyi").read_text())
    surface = set()
    for node in tree.body:
        if not isinstance(node, ast.ClassDef):
            continue
        for f in node.body:
            if not isinstance(f, ast.FunctionDef):
                continue
            decorators = [ast.unparse(d) for d in f.decorator_list]
            setter_ = any(d.endswith(".setter") for d in decorators)
            if f.name in ("__setitem__", "__delitem__") or setter_ or (
                    not f.name.startswith("_") and not set(decorators) & {"property", "staticmethod", "classmethod"}
                    and not f.name.startswith(("render", "to_"))):
                surface.add(f"{node.name}.{f.name}")
    return surface - READS


class Deck:
    """One deck opened from `path`, with a probe slide that each timed call adds its targets to; reopened every
    REOPEN calls."""

    def __init__(self, path):
        self.path, self.uses, self.prs, self.slide = path, REOPEN, None, None

    def probe(self):
        if self.uses == REOPEN:
            self.prs = rpptx.Presentation(self.path)
            self.slide, self.uses = self.prs.slides.add_slide(self.prs.slide_layouts[6]), 0
        self.uses += 1
        return self.prs, self.slide


def costs_ms(decks, setup):
    """The fastest of CALLS calls on fresh targets in each of `decks`, or of two calls once each one takes over 50 ms.
    The decks take turns, so that a change of the machine's load hits all of them."""
    readies = [setup(*deck.probe()) for deck in decks]
    best = [float("inf")] * len(decks)
    for k in range(CALLS):
        for i, ready in enumerate(readies):
            action = ready(k)
            start = time.perf_counter()
            action()
            best[i] = min(best[i], time.perf_counter() - start)
        if k and min(best) > 0.05:
            break
    return [b * 1000 for b in best]


@pytest.fixture(scope="module")
def decks(tmp_path_factory):
    folder = tmp_path_factory.mktemp("scaling")
    return (Deck(edit_cost_pptx(folder / "n.pptx", N, picture_side=600)),
            Deck(edit_cost_pptx(folder / "4n.pptx", 4 * N, picture_side=600)))


# Calls whose cost grows with the deck on the pin, in the order of TIMED: none since tensorbee/rdocx#413 and #416
# fixed the 159 of rpptx 0.14.3 (the last nine, slide transitions and header and footer setters, flat on every CI
# platform with integration/open-prs-2026-10-23-5). A call that grows again is a regression: it comes back here only
# with an entry for GROWS in tests/gaps.py and an upstream issue.
GROWING = set()


@pytest.mark.parametrize("name", [pytest.param(name, marks=pytest.mark.gap(GROWS))
                                  if name in GROWING else name for name in TIMED])
def test_call_cost_does_not_grow_with_the_deck(decks, name):
    small, large = costs_ms(decks, TIMED[name])
    bound = (FIXED if name in GROWING else RATIO) * small + FLOOR_MS
    assert large < bound, f"{name}: {large:.1f} ms at {4 * N} slides, {small:.1f} ms at {N}"


def test_every_mutating_call_is_timed():
    assert sorted(GROWING - set(TIMED)) == []
    timed = {name.split(" (")[0] for name in TIMED}
    assert sorted(mutating_surface() - timed - set(WHOLE_DECK)) == []
    assert sorted((timed | set(WHOLE_DECK)) - mutating_surface()) == []
