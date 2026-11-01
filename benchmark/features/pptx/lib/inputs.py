"""Deterministic input decks for the pptx feature tasks (used by every make_input.py, never by an implementation).

The decks are built with python-pptx from its default template, then given what a real corporate deck has:
16:9 masters and layouts, a named theme with its own colours and fonts (Cambria headings, Calibri body, rendered
with Caladea and Carlito), a logo on the master, placeholders, runs with mixed formatting, notes, groups, tables,
native charts and pictures. Every zip is rewritten with fixed timestamps so the bytes are reproducible.
"""

from __future__ import annotations

import copy
import io
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Inches, Pt  # noqa: F401  (re-exported for make_input.py)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

NORTHWIND = {
    "name": "Northwind",
    "major": "Cambria",
    "minor": "Calibri",
    "colors": {
        "dk1": "1F1F1F", "lt1": "FFFFFF", "dk2": "1E2761", "lt2": "EEF1F8",
        "accent1": "1E2761", "accent2": "C8553D", "accent3": "2A9D8F", "accent4": "E9C46A",
        "accent5": "6D597A", "accent6": "8AB17D", "hlink": "2A6FDB", "folHlink": "7B4EA3",
    },
}
HARBOR = {
    "name": "Harbor",
    "major": "Calibri",
    "minor": "Cambria",
    "colors": {
        "dk1": "102A43", "lt1": "FFFFFF", "dk2": "0B6E4F", "lt2": "F0F4F0",
        "accent1": "0B6E4F", "accent2": "F18F01", "accent3": "C73E1D", "accent4": "3B8EA5",
        "accent5": "8E7DBE", "accent6": "A5BE00", "hlink": "005F73", "folHlink": "9B2226",
    },
}

LAYOUT = {  # python-pptx default template layout indexes
    "title": 0, "content": 1, "section": 2, "two": 3, "comparison": 4, "title_only": 5, "blank": 6,
    "content_caption": 7, "picture": 8,
}


# --- pictures ----------------------------------------------------------------------------------
def png(width=640, height=400, seed=0, label=None) -> bytes:
    """A deterministic PNG: a two-colour gradient with a few shapes (stands for a photo or a figure)."""
    palettes = [((30, 39, 97), (200, 85, 61)), ((11, 110, 79), (241, 143, 1)), ((109, 89, 122), (42, 157, 143)),
                ((233, 196, 106), (31, 31, 31))]
    c0, c1 = palettes[seed % len(palettes)]
    img = Image.new("RGB", (width, height))
    d = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(1, height - 1)
        d.line([(0, y), (width, y)], fill=tuple(int(c0[k] + (c1[k] - c0[k]) * t) for k in range(3)))
    for k in range(4):
        x = (seed * 53 + k * 157) % (width - 80)
        y = (seed * 31 + k * 97) % (height - 80)
        d.ellipse([x, y, x + 60 + 10 * k, y + 60 + 10 * k], fill=(255, 255, 255) if k % 2 else c1, outline=(0, 0, 0))
    if label:
        d.rectangle([10, height - 40, 10 + 12 * len(label), height - 10], fill=(255, 255, 255))
        d.text((16, height - 34), label, fill=(0, 0, 0))
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


# --- the base deck -----------------------------------------------------------------------------
def _theme_part(prs):
    return prs.slide_master.part.part_related_by(RT.THEME)


def set_theme(prs, theme) -> None:
    part = _theme_part(prs)
    root = etree.fromstring(part.blob)
    root.set("name", theme["name"])
    scheme = root.find(".//a:clrScheme", NS)
    scheme.set("name", theme["name"])
    for key, val in theme["colors"].items():
        slot = scheme.find(f"a:{key}", NS)
        for child in list(slot):
            slot.remove(child)
        etree.SubElement(slot, f"{{{A}}}srgbClr", val=val)
    fonts = root.find(".//a:fontScheme", NS)
    fonts.set("name", theme["name"])
    fonts.find("a:majorFont/a:latin", NS).set("typeface", theme["major"])
    fonts.find("a:minorFont/a:latin", NS).set("typeface", theme["minor"])
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _widen(element, factor) -> None:
    for off in element.iter(f"{{{A}}}off", f"{{{A}}}chOff"):
        off.set("x", str(int(int(off.get("x")) * factor)))
    for ext in element.iter(f"{{{A}}}ext", f"{{{A}}}chExt"):
        if ext.get("cx") is not None:
            ext.set("cx", str(int(int(ext.get("cx")) * factor)))


def base_deck(theme=NORTHWIND, logo=True, layout_names=None):
    """python-pptx's default master and eleven layouts, widened to 16:9, with `theme` and a logo on the master."""
    prs = Presentation()
    factor = 12192000 / prs.slide_width
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    _widen(prs.slide_master.element, factor)
    for layout in prs.slide_layouts:
        _widen(layout.element, factor)
    set_theme(prs, theme)
    if layout_names:
        for layout, name in zip(prs.slide_layouts, layout_names):
            layout.element.find("p:cSld", NS).set("name", name)
    if logo:  # a small logo block on the master: on every slide, not slide text
        from pptx.oxml.shapes.autoshape import CT_Shape
        tree = prs.slide_master.shapes._spTree
        sp = CT_Shape.new_autoshape_sp(90, "Logo", "roundRect", Inches(12.35), Inches(0.2), Inches(0.75), Inches(0.4))
        tree.append(sp)
        m = list(prs.slide_master.shapes)[-1]
        m.fill.solid()
        m.fill.fore_color.rgb = RGBColor.from_string(theme["colors"]["accent2"])
        m.line.fill.background()
        m.text_frame.text = theme["name"][:2].upper()
        m.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
        m.text_frame.paragraphs[0].runs[0].font.bold = True
    cp = prs.core_properties
    cp.title, cp.author, cp.last_modified_by = f"{theme['name']} deck", "Northwind Strategy Team", "Northwind Strategy Team"
    import datetime as _dt
    cp.created = cp.modified = _dt.datetime(2026, 9, 1, 9, 0, 0)
    cp.revision = 3
    return prs


def layout(prs, key):
    return prs.slide_layouts[LAYOUT[key]]


# --- text helpers ------------------------------------------------------------------------------
def runs(paragraph, parts) -> None:
    """parts: [text | (text, {bold, italic, underline, color, size})]; the paragraph's existing runs are dropped."""
    for r in list(paragraph.runs):
        r._r.getparent().remove(r._r)
    for part in parts:
        text, fmt = (part, {}) if isinstance(part, str) else part
        r = paragraph.add_run()
        r.text = text
        f = r.font
        if fmt.get("bold"):
            f.bold = True
        if fmt.get("italic"):
            f.italic = True
        if fmt.get("underline"):
            f.underline = True
        if fmt.get("color"):
            f.color.rgb = RGBColor.from_string(fmt["color"])
        if fmt.get("size"):
            f.size = Pt(fmt["size"])


def bullets(text_frame, items) -> None:
    """items: [(level, parts)], parts as in runs()."""
    for k, (level, parts) in enumerate(items):
        p = text_frame.paragraphs[0] if k == 0 else text_frame.add_paragraph()
        p.level = level
        runs(p, parts)


def textbox(shapes, x, y, w, h, parts, size=None, name=None, wrap=True):
    tb = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.text_frame.word_wrap = wrap
    if name:
        tb.name = name
    runs(tb.text_frame.paragraphs[0], parts)
    if size:
        for r in tb.text_frame.paragraphs[0].runs:
            if r.font.size is None:
                r.font.size = Pt(size)
    return tb


def table(shapes, x, y, w, h, rows, name="Table", first_row=True, fills=None, size=14):
    gf = shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    gf.name = name
    t = gf.table
    t.first_row = first_row
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
                    if r == 0:
                        run.font.bold = True
    for (r, c), color in (fills or {}).items():
        t.cell(r, c).fill.solid()
        t.cell(r, c).fill.fore_color.rgb = RGBColor.from_string(color)
    return gf


def chart(shapes, x, y, w, h, categories, series, name="Chart", title=None, colors=None, labels=True,
          number_format='#,##0'):
    data = CategoryChartData(number_format=number_format)
    data.categories = categories
    for sname, values in series:
        data.add_series(sname, values)
    gf = shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data)
    gf.name = name
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    if title:
        ch.has_title = True
        ch.chart_title.text_frame.text = title
    for k, s in enumerate(ch.series):
        if colors:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = RGBColor.from_string(colors[k % len(colors)])
    if labels:
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.number_format = number_format
        plot.data_labels.number_format_is_linked = False
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        plot.data_labels.font.size = Pt(10)
    return gf


def notes(slide, text) -> None:
    slide.notes_slide.notes_text_frame.text = text


# --- saving ------------------------------------------------------------------------------------
FIXED = (2026, 9, 1, 9, 0, 0)


def normalize_zip(path, content_type_patch=None) -> None:
    """Rewrite the zip with fixed timestamps (and optionally patch [Content_Types].xml) so bytes are reproducible."""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items:
            if name == "[Content_Types].xml" and content_type_patch:
                data = content_type_patch(data)
            info = zipfile.ZipInfo(name, FIXED)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            z.writestr(info, data)


def save(prs, path, content_type_patch=None) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    normalize_zip(path, content_type_patch)
    return path


def to_template(path_pptx, path_potx) -> Path:
    """Copy a .pptx as a .potx (template main content type)."""
    path_potx = Path(path_potx)
    data = Path(path_pptx).read_bytes()
    path_potx.write_bytes(data)
    normalize_zip(path_potx, lambda b: b.replace(b"presentationml.presentation.main+xml",
                                                  b"presentationml.template.main+xml"))
    return path_potx


# --- the standard review deck (P02, P12, P13) --------------------------------------------------
REVIEW = {
    "titles": ["Alpha: Northwind quarterly review", "Bravo: Agenda", "Charlie: Market view", "Delta: Regional results",
               "Echo: Delivery process", "Foxtrot: Revenue trend", "Golf: New headquarters", "Hotel: Questions"],
    "notes": {1: "mk-note-1 Welcome everyone and set the scene.", 2: "mk-note-2 Keep the agenda short.",
              3: "mk-note-3 Market view comes from the September survey.", 4: "mk-note-4 North beat plan by 8 %.",
              5: "mk-note-5 Walk through the three stages.", 6: "mk-note-6 Revenue grew every quarter.",
              8: "mk-note-8 Leave ten minutes for questions."},
    # marker strings that must be found in each slide's text (titles excluded)
    "markers": {1: ["mk-sub-1"], 2: ["mk-body-2", "mk-sub-bullet-2"], 3: ["mk-left-3", "mk-right-3"],
                4: ["mk-cell-4"], 5: ["mk-group-5", "mk-group-5b"], 6: ["mk-caption-6"], 7: ["mk-caption-7"],
                8: ["mk-body-8"]},
}


def review_deck(photo: bytes):
    prs = base_deck()
    T = REVIEW["titles"]
    s = prs.slides.add_slide(layout(prs, "title"))
    s.shapes.title.text = T[0]
    runs(s.placeholders[1].text_frame.paragraphs[0], ["Q3 2026 ", ("|", {"color": "C8553D"}), " Strategy team mk-sub-1"])

    s = prs.slides.add_slide(layout(prs, "content"))
    s.shapes.title.text = T[1]
    bullets(s.placeholders[1].text_frame, [
        (0, ["Results ", ("by region", {"bold": True}), " mk-body-2"]),
        (1, [("Detail: ", {"italic": True}), "North and South mk-sub-bullet-2"]),
        (0, ["Revenue ", ("trend", {"color": "2A9D8F", "bold": True})]),
        (0, ["Delivery process"]),
        (0, ["Questions"]),
    ])

    s = prs.slides.add_slide(layout(prs, "two"))
    s.shapes.title.text = T[2]
    bullets(s.placeholders[1].text_frame, [(0, ["Demand is ", ("rising", {"bold": True}), " mk-left-3"]),
                                           (0, ["Prices are stable"])])
    bullets(s.placeholders[2].text_frame, [(0, [("Competition", {"italic": True}), " is consolidating mk-right-3"]),
                                           (0, ["Two new entrants"])])

    s = prs.slides.add_slide(layout(prs, "title_only"))
    s.shapes.title.text = T[3]
    table(s.shapes, 0.8, 1.7, 11.7, 3.2, [["Region", "Revenue", "Plan", "Delta"],
                                          ["North", "1,240", "1,150", "+8 %"],
                                          ["South", "980", "1,000", "-2 % mk-cell-4"],
                                          ["West", "760", "700", "+9 %"]], name="Results table")

    s = prs.slides.add_slide(layout(prs, "title_only"))
    s.shapes.title.text = T[4]
    g = s.shapes.add_group_shape()
    g.name = "Process"
    for k, (label, extra) in enumerate([("Plan", ""), ("Build", " mk-group-5"), ("Run", " mk-group-5b")]):
        box = g.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + 4 * k), Inches(2.5), Inches(3.4), Inches(1.6))
        box.name = f"Step {k + 1}"
        box.text_frame.text = label + extra

    s = prs.slides.add_slide(layout(prs, "title_only"))
    s.shapes.title.text = T[5]
    chart(s.shapes, 0.8, 1.6, 8.5, 4.8, ["Q1", "Q2", "Q3", "Q4"], [("Revenue", (820, 910, 1015, 1120))],
          name="Revenue chart", title="Revenue (EUR k)", colors=["1E2761"])
    textbox(s.shapes, 9.6, 2.0, 3.0, 2.0, ["Revenue grew ", ("every quarter", {"bold": True}), " mk-caption-6"], size=16,
            name="Caption")

    s = prs.slides.add_slide(layout(prs, "picture"))
    s.shapes.title.text = T[6]
    s.placeholders[1].insert_picture(io.BytesIO(photo))
    runs(s.placeholders[2].text_frame.paragraphs[0], ["Opening in March mk-caption-7"])

    s = prs.slides.add_slide(layout(prs, "section"))
    s.shapes.title.text = T[7]
    runs(s.placeholders[1].text_frame.paragraphs[0], ["Thank you mk-body-8"])

    for n, text in REVIEW["notes"].items():
        notes(prs.slides[n - 1], text)
    return prs
