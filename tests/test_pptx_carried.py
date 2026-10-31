"""pptx: fixes the pinned integration build (upstream rpptx-v0.14.4 plus tensorbee/rdocx#438, #439 and #442) carries
and upstream `main` does not have yet, which the skills rely on, and the cases of those issues that stay open as known
gaps (strict xfails, tests/gaps.py). Each test names the pull request that carries it, or the issue of its gap: delete
the test once upstream main has it. Each test builds its own input."""
import collections
import io
import re
import struct
import warnings
import xml.etree.ElementTree as ET
import zipfile
import zlib

import pytest
import rpptx
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches

from conftest import poppler, run


# ---- tensorbee/rdocx#439 (#435) and #438 (#431) ----------------------------------------------------------------------


def rezip(data, edits=None, duplicate=None):
    """The package data with the parts named in edits (name: function of the part's bytes) replaced and, when
    duplicate names a part, a second entry of that name, its "First copy" changed to "Second copy"."""
    edits = edits or {}
    zin, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            x = zin.read(item)
            z.writestr(item, edits[item.filename](x) if item.filename in edits else x)
            if item.filename == duplicate:
                with pytest.warns(UserWarning, match="Duplicate name"):
                    z.writestr(item.filename, x.replace(b"First copy", b"Second copy"))
    return out.getvalue()


def validate(cli, path, data):
    """Exit code and output of `rpptx validate` on data written to path."""
    path.write_bytes(data)
    res = run([cli, "validate", path])
    return res.returncode, res.stdout + res.stderr


def well_formed(xml):
    try:
        ET.fromstring(xml)
    except ET.ParseError:
        return False
    return True


def text_edit(edit):
    """edit, a function of a part's text, as a function of its bytes that must change them."""
    def on_bytes(x):
        new = edit(x.decode())
        assert new != x.decode()
        return new.encode()
    return on_bytes


def unescaped_company(x):
    """docProps/app.xml with the Company property set to 'Smith & Sons', the ampersand left unescaped."""
    x = re.sub(rb"<Company>[^<]*</Company>|<Company/>", b"", x)
    return x.replace(b"</Properties>", b"<Company>Smith & Sons</Company></Properties>")


def first_copy_pptx():
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "First copy"
    return prs.to_bytes()


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_a_malformed_properties_part_and_a_part_stored_twice(rpptx_cli, tmp_path):
    """`rpptx validate` exits non-zero naming the part on a docProps/app.xml that is not well-formed and on a ZIP
    that stores ppt/slides/slide1.xml twice (cli.md, validate). It used to exit 0 on both (tensorbee/rdocx#435)."""
    path = tmp_path / "case.pptx"
    code, out = validate(rpptx_cli, path, rezip(first_copy_pptx(), {"docProps/app.xml": unescaped_company}))
    assert code != 0 and "app.xml" in out
    code, out = validate(rpptx_cli, path, rezip(first_copy_pptx(), duplicate="ppt/slides/slide1.xml"))
    assert code != 0 and "ppt/slides/slide1.xml" in out


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_open_refuses_or_warns_on_a_part_stored_twice(tmp_path):
    """rpptx.Presentation() refuses, or warns about, a ZIP that stores ppt/slides/slide1.xml twice, naming the part.
    It used to open it silently and read the last copy, where other readers read the first (tensorbee/rdocx#435)."""
    path = tmp_path / "case.pptx"
    path.write_bytes(rezip(first_copy_pptx(), duplicate="ppt/slides/slide1.xml"))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            rpptx.Presentation(str(path))
            refusal = ""
        except Exception as e:
            refusal = str(e)
    assert "ppt/slides/slide1.xml" in refusal or any("ppt/slides/slide1.xml" in str(w.message) for w in caught)


SLIDE, PRES, TYPES = "ppt/slides/slide1.xml", "ppt/presentation.xml", "[Content_Types].xml"
STORED_FAULTS = {  # label: (part, edit of its text, text the output must hold)
    "unknown element p:bogus in p:spTree": (
        SLIDE, lambda x: x.replace("</p:spTree>", "<p:bogus/></p:spTree>", 1), "slide1.xml"),
    "p:cNvPr without its required name": (
        SLIDE, lambda x: re.sub(r'(<p:cNvPr id="\d+") name="Title 1"', r"\1", x, count=1), "slide1.xml"),
    "p:notesMasterIdLst after p:sldIdLst": (PRES, lambda x: re.sub(
        r"(<p:notesMasterIdLst>.*?</p:notesMasterIdLst>)(<p:sldIdLst>.*?</p:sldIdLst>)", r"\2\1", x, count=1,
        flags=re.S), "presentation.xml"),
    "p:sldSz before p:sldIdLst": (PRES, lambda x: re.sub(
        r"(<p:sldIdLst>.*?</p:sldIdLst>)(<p:sldSz [^>]*/>)", r"\2\1", x, count=1, flags=re.S), "presentation.xml"),
    "slide Override removed from [Content_Types].xml": (TYPES, lambda x: re.sub(
        r'\s*<Override PartName="/ppt/slides/slide1.xml"[^>]*/>', "", x, count=1), "slide1.xml"),
    "table row with 1 a:tc for 2 a:gridCol": (SLIDE, lambda x: re.sub(
        r"(<a:tr [^>]*>)(<a:tc>.*?</a:tc>)(<a:tc>.*?</a:tc>)(</a:tr>)", r"\1\2\4", x, count=1, flags=re.S),
        "slide1.xml"),
}


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_stored_presentationml_faults(rpptx_cli, tmp_path):
    """`rpptx validate` exits non-zero naming the part on an unknown shape tree child, a p:cNvPr without name,
    misordered presentation.xml children, a slide without its content type and a table row short of a cell, and
    passes the deck they are made from (cli.md, validate). It used to exit 0 on all six, since it checked the deck
    as parsed, not the stored parts (tensorbee/rdocx#435)."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Hello"
    slide.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1))
    slide.notes_slide.notes_text_frame.text = "A note"
    base = prs.to_bytes()
    path = tmp_path / "case.pptx"
    assert validate(rpptx_cli, path, base)[0] == 0
    missed = {}
    for label, (part, edit, needle) in STORED_FAULTS.items():
        code, out = validate(rpptx_cli, path, rezip(base, {part: text_edit(edit)}))
        if code == 0 or needle not in out:
            missed[label] = (code, out[:200])
    assert not missed


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_counts_the_grid_span_of_a_table_row(rpptx_cli, tmp_path):
    """`rpptx validate` counts each a:tc's gridSpan against the a:gridCol count: a row whose first cell spans both
    columns, with no hMerge cell after it, as LibreOffice writes one, passes; a row of one plain cell for two columns
    is still refused, naming what it covers (cli.md, validate). The first version of #439 refused the spanned row."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Hello"
    slide.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1))
    base = prs.to_bytes()
    path = tmp_path / "case.pptx"
    spanned = text_edit(lambda x: re.sub(r"(<a:tr [^>]*>)<a:tc>(.*?</a:tc>)<a:tc>.*?</a:tc>(</a:tr>)",
                                         r'\1<a:tc gridSpan="2">\2\3', x, count=1, flags=re.S))
    code, out = validate(rpptx_cli, path, rezip(base, {SLIDE: spanned}))
    assert code == 0, out
    short = text_edit(STORED_FAULTS["table row with 1 a:tc for 2 a:gridCol"][1])
    code, out = validate(rpptx_cli, path, rezip(base, {SLIDE: short}))
    assert code != 0 and "covers 1 of its 2 a:gridCol" in out, out


def control_character_pptx(path):
    """A deck whose ppt/slides/slide1.xml holds U+0001 inside its title, written to path."""
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "Hello"
    path.write_bytes(rezip(prs.to_bytes(), {SLIDE: lambda x: x.replace(b">Hello<", b">He\x01llo<")}))
    assert not well_formed(zipfile.ZipFile(path).read(SLIDE))
    return path


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
def test_validate_reports_a_part_holding_a_control_character(rpptx_cli, tmp_path):
    """`rpptx validate` exits non-zero naming slide1.xml when it holds U+0001, which XML 1.0 forbids (cli.md,
    validate). It used to exit 0 (tensorbee/rdocx#435)."""
    path = control_character_pptx(tmp_path / "ctrl.pptx")
    res = run([rpptx_cli, "validate", path])
    assert res.returncode != 0 and "slide1.xml" in res.stdout + res.stderr


# upstream: tensorbee/rdocx#435
@pytest.mark.gap("pptx-control-character-round-trip")
def test_open_refuses_a_control_character_or_save_writes_a_well_formed_part(tmp_path):
    """A deck whose ppt/slides/slide1.xml holds U+0001 is refused at open, naming the part, or opens and save()
    writes a well-formed part. Today it opens and save() writes the U+0001 back raw (tensorbee/rdocx#435)."""
    path = control_character_pptx(tmp_path / "ctrl.pptx")
    try:
        prs = rpptx.Presentation(str(path))
    except Exception as e:
        assert "slide1.xml" in str(e)
        return
    prs.save(str(tmp_path / "saved.pptx"))
    assert well_formed(zipfile.ZipFile(tmp_path / "saved.pptx").read(SLIDE))


NEGATIVE_SIZES = {
    "add_textbox(width=-3 in)": lambda s: s.shapes.add_textbox(Inches(1), Inches(1), -Inches(3), Inches(1)),
    "add_shape(height=-1 in)": lambda s: s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3),
                                                            -Inches(1)),
    "add_table(width=-3 in)": lambda s: s.shapes.add_table(2, 2, Inches(1), Inches(1), -Inches(3), Inches(1)),
}


# carried by tensorbee/rdocx#439 (#435), delete once upstream main has it
@pytest.mark.parametrize("call", NEGATIVE_SIZES)
def test_new_shapes_refuse_a_negative_size(call):
    """add_textbox and add_shape refuse a negative width or height at the call, as add_table does. They used to
    accept it and save a negative a:ext that rpptx could not open again (tensorbee/rdocx#435)."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    with pytest.raises(Exception):
        NEGATIVE_SIZES[call](slide)


SLIDE_NUMBER = ('<p:sp><p:nvSpPr><p:cNvPr id="9" name="Slide Number Placeholder 3"/><p:cNvSpPr/><p:nvPr>'
                '<p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>'
                '<a:lstStyle/><a:p><a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"><a:t>1</a:t>'
                '</a:fld></a:p></p:txBody></p:sp>')


def duplicate_shape_id(x):
    """slide1.xml with its second text box given the id of the first."""
    ids = re.findall(r'<p:cNvPr id="(\d+)"', x)  # group, first box, second box
    return x.replace(f'<p:cNvPr id="{ids[2]}"', f'<p:cNvPr id="{ids[1]}"', 1)


UNASSESSABLE = {  # label: (part, edit of its text)
    "duplicate shape id": (SLIDE, duplicate_shape_id),
    "notes placeholder idx 10": ("ppt/notesSlides/notesSlide1.xml",
                                 lambda x: x.replace("</p:spTree>", SLIDE_NUMBER + "</p:spTree>")),
}


def unassessable_deck(path, fault):
    """A one-slide deck (two text boxes and a note) with the fault UNASSESSABLE[fault], written to path."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text_frame.text = "First"
    slide.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1)).text_frame.text = "Second"
    slide.notes_slide.notes_text_frame.text = "A note"
    part, edit = UNASSESSABLE[fault]
    path.write_bytes(rezip(prs.to_bytes(), {part: text_edit(edit)}))
    return path


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
@pytest.mark.parametrize("fault", UNASSESSABLE)
def test_fit_reports_a_deck_holding_a_frame_it_cannot_assess(rpptx_cli, tmp_path, fault):
    """`rpptx fit` gives its report (exit 0 or 1) and fit_report() resolves every frame of a deck that renders and
    holds a repeated shape id or a notes placeholder whose idx the notes master lacks (cli.md and python-api.md,
    fit). `rpptx fit` used to exit 2 with no report, and fit_report() raised or listed an "uninstantiated" frame
    (tensorbee/rdocx#431)."""
    path = unassessable_deck(tmp_path / "deck.pptx", fault)
    assert run([rpptx_cli, "render", path, "-o", tmp_path / "pages", "--dpi", "20", "--force"]).returncode == 0
    assert run([rpptx_cli, "fit", path]).returncode in (0, 1)
    frames = rpptx.Presentation(str(path)).fit_report().frames
    assert not [(f.context, f.overflow) for f in frames if "uninstantiated" in f.context]


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
def test_validate_names_a_repeated_shape_id_in_a_sentence(rpptx_cli, tmp_path):
    """`rpptx validate` reports a repeated shape id as a sentence. It used to print a Rust debug value
    ("DuplicateShapeId { slide: 0, id: 2 }", tensorbee/rdocx#431)."""
    res = run([rpptx_cli, "validate", unassessable_deck(tmp_path / "deck.pptx", "duplicate shape id")])
    line = next((l.strip() for l in (res.stderr + res.stdout).splitlines() if "uplicate" in l), "")
    assert line and "{" not in line


def table_deck(top, repeat):
    """A title-only slide with a 2 x 2 table at top inches, rows stored 0.4 in high, its last cell repeat sentences."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Risks"
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(top), Inches(10), Inches(0.8)).table
    for r in range(2):
        table.rows[r].height = Inches(0.4)
    table.cell(0, 0).text, table.cell(0, 1).text, table.cell(1, 0).text = "Risk", "Mitigation", "Hiring"
    table.cell(1, 1).text = "Start interviews in November with two agencies; " * repeat
    return prs


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
def test_fit_report_does_not_overflow_a_table_row_that_grows_on_the_slide():
    """fit_report() gives no overflow (or a status distinct from a table that runs off the slide) for a cell taller
    than its stored row when the row grows and the table stays on the slide, as in the render; a table that runs off
    the slide still overflows (python-api.md, fit_report). The first used to be an overflow needing a 40 % font
    scale (tensorbee/rdocx#431)."""
    from PIL import Image

    dpi, seen = 40, {}
    for label, top, repeat in (("grows", 2.0, 4), ("runs off", 6.0, 8)):
        prs = table_deck(top, repeat)
        frame = next(f for f in prs.fit_report().frames if f.update.cell == (1, 1))
        image = Image.open(io.BytesIO(prs.render_slide_to_png(0, dpi))).convert("RGB")
        x0, x1 = int(1.2 * dpi), int(10.8 * dpi)
        bottom = max(y for y in range(image.height)
                     if min(min(image.getpixel((x, y))) for x in range(x0, x1)) < 250) / dpi
        seen[label] = frame, bottom, image.height / dpi
    (grows, grows_bottom, slide_height), (off, _, _) = seen["grows"], seen["runs off"]
    assert grows_bottom < slide_height - 0.5
    assert not grows.overflow or grows.update.status != off.update.status
    assert off.overflow


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
def test_fit_text_output_names_each_frame_once(rpptx_cli, tmp_path):
    """The plain-text output of `rpptx fit` gives each frame its own line, with its child path in plain indices and
    the shape name of each table cell (cli.md, fit). It used to name frames by a Rust debug value
    ("Array [Number(1)]"), and 33 of 52 lines repeated another (tensorbee/rdocx#431)."""
    prs = rpptx.Presentation()
    for i in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Slide {i + 1}"
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(8), Inches(1)).table
    for r in range(2):
        for c in range(2):
            table.cell(r, c).text = f"r{r}c{c}"
    path = tmp_path / "fit-text.pptx"
    prs.save(str(path))
    out = run([rpptx_cli, "fit", path], check=True).stdout
    lines = [l for l in out.splitlines() if l.strip() and not l.startswith("Diagnostic")]
    assert sum(n - 1 for n in collections.Counter(lines).values() if n > 1) == 0
    assert not [l for l in lines if "Array [" in l or "Number(" in l]
    assert len([l for l in lines if "slide3.xml" in l and "Table 3" in l]) == 4


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
def test_markdown_export_writes_plain_spaces_and_escapes_only_markdown_syntax():
    """export_markdown() writes a title, a bullet and the notes with plain spaces and no escape of punctuation that
    Markdown would not read as syntax, and keeps a repeated space with &#32; (python-api.md, export_markdown). It
    used to write every space as &#32; and escape "(", ")", "." and "-" mid-sentence (tensorbee/rdocx#431)."""
    title, body = "Q3 revenue, by region", "North grew 12% (EUR 1.2m) - see table."
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    slide.placeholders[1].text = body
    slide.notes_slide.notes_text_frame.text = "Thank the team."
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "Two  spaces"
    lines = prs.export_markdown(include_notes=True).text.splitlines()

    def line(prefix):
        return next((l for l in lines if l.startswith(prefix)), None)

    assert line("# Q3") == "# " + title
    assert line("- North") == "- " + body
    assert line("> Thank") == "> Thank the team."
    assert "&#32;" in (line("# Two") or "")


# carried by tensorbee/rdocx#438 (#431), delete once upstream main has it
def test_slides_index_returns_the_position_of_a_slide():
    """prs.slides.index(slide) returns the slide's zero-based position, the value slides.move() works with, and
    raises ValueError for a slide of another presentation, as python-pptx's Slides.index. It used to raise a bare
    AttributeError (tensorbee/rdocx#431)."""
    prs = rpptx.Presentation()
    for title in ("One", "Two", "Three"):
        prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = title
    other = rpptx.Presentation()
    foreign = other.slides.add_slide(other.slide_layouts[5])
    assert prs.slides.index(prs.slides[2]) == 2
    with pytest.raises(ValueError):
        prs.slides.index(foreign)


# ---- tensorbee/rdocx#442 (#436) --------------------------------------------------------------------------------------


RED = (200, 30, 30)


def decode_png(data):
    """Rows of (r, g, b) of the 8-bit RGB or RGBA, non-interlaced PNG the renderer writes."""
    pos, idat = 8, b""
    while pos < len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        body = data[pos + 8:pos + 8 + length]
        if kind == b"IHDR":
            width, height, depth, ctype = struct.unpack(">IIBB", body[:10])
            assert depth == 8 and ctype in (2, 6), (depth, ctype)
            bpp = 3 if ctype == 2 else 4
        elif kind == b"IDAT":
            idat += body
        pos += 12 + length
    raw, stride, rows, prev = zlib.decompress(idat), width * bpp, [], bytearray(width * bpp)
    for y in range(height):
        f, line = raw[y * (stride + 1)], bytearray(raw[y * (stride + 1) + 1:(y + 1) * (stride + 1)])
        for i in range(stride):
            a, b, c = (line[i - bpp], prev[i], prev[i - bpp]) if i >= bpp else (0, prev[i], 0)
            p = a + b - c
            paeth = a if abs(p - a) <= abs(p - b) and abs(p - a) <= abs(p - c) else b if abs(p - b) <= abs(p - c) else c
            line[i] = (line[i] + (0, a, b, (a + b) // 2, paeth)[f]) & 255
        rows.append([tuple(line[x * bpp:x * bpp + 3]) for x in range(width)])
        prev = line
    return rows


def colours(png):
    """(red pixels, pixels of another saturated colour) of a rendered PNG."""
    red = other = 0
    for row in decode_png(png):
        for r, g, b in row:
            if r > 150 and g < 80 and b < 80:
                red += 1
            elif max(r, g, b) - min(r, g, b) > 80:
                other += 1
    return red, other


def red_emf(w=200, h=100):
    """A minimal EMF: header, a solid red brush, one filled rectangle, end of file ([MS-EMF] 2.3)."""
    records = [
        struct.pack("<IIIIII", 39, 24, 1, 0, 0x001E1EC8, 0),  # EMR_CREATEBRUSHINDIRECT: handle 1, solid, red
        struct.pack("<III", 37, 12, 1),  # EMR_SELECTOBJECT: handle 1
        struct.pack("<II4i", 43, 24, 10, 10, w - 10, h - 10),  # EMR_RECTANGLE
        struct.pack("<IIIII", 14, 20, 0, 16, 20),  # EMR_EOF
    ]
    body = b"".join(records)
    header = struct.pack("<II4i4iIIIIHHIIIiiii", 1, 88, 0, 0, w - 1, h - 1, 0, 0, w * 2645 // 100, h * 2645 // 100,
                         0x464D4520, 0x10000, 88 + len(body), 1 + len(records), 2, 0, 0, 0, 0, 1024, 768, 270, 203)
    return header + body


def red_png(interlaced, w=64, h=48):
    """A uniformly red RGB PNG, plain or interlaced (Adam7: seven passes, each with its own filter bytes)."""
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    passes = ((0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2))
    raw = b"".join((b"\x00" + bytes(RED) * len(range(x0, w, dx))) * len(range(y0, h, dy))
                   for x0, y0, dx, dy in (passes if interlaced else ((0, 0, 1, 1),)) if x0 < w and y0 < h)
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 1 if interlaced else 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def category_data(categories, series, values, **kw):
    data = rpptx.CategoryChartData(**kw)
    data.categories = categories
    data.add_series(series, values)
    return data


def chart_deck(kind, data, layout=6):
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[layout])
    slide.shapes.add_chart(kind, Inches(1), Inches(1.5), Inches(6), Inches(4), data)
    return prs


def picture_render(picture, *size):
    prs = rpptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(io.BytesIO(picture), Inches(1), Inches(1), *size)
    return prs.render_slide_to_png_with_diagnostics(0, 30)


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_render_draws_an_emf_picture():
    """An EMF picture (a red filled rectangle) is drawn in the slide render. It used to be left out, with the
    diagnostic "missing slide picture image relationship" although the relationship and the part were there
    (tensorbee/rdocx#436)."""
    png, _ = picture_render(red_emf(), Inches(3), Inches(1.5))
    assert colours(png)[0] >= 2700  # two thirds of the 3 x 1.5 in frame at 30 dpi (4050 pixels)


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_render_draws_an_interlaced_png_like_its_plain_copy():
    """An interlaced (Adam7) PNG renders like its plain copy: the same red pixels, no other colour, no diagnostic. It
    used to be left out, with the diagnostic "missing slide picture image relationship" (tensorbee/rdocx#436)."""
    plain_red = colours(picture_render(red_png(False), Inches(3))[0])[0]
    png, diagnostics = picture_render(red_png(True), Inches(3))
    assert colours(png) == (plain_red, 0)
    assert not list(diagnostics)


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
@pytest.mark.parametrize("kind", ["COLUMN_CLUSTERED", "LINE", "AREA"])
def test_a_chart_with_a_third_axis_id_is_read_and_drawn(kind):
    """A bar, line or area chart whose plot lists a third c:axId naming no axis, as some deck generators write it, is
    read and drawn, the extra id ignored. Reading it used to raise ("ChartML contains duplicate c:axId") and the
    render drew an "Unsupported chart" placeholder (tensorbee/rdocx#436)."""
    package = chart_deck(getattr(rpptx.XL_CHART_TYPE, kind), category_data(["x", "y"], "A", (1, 2))).to_bytes()
    archive, out = zipfile.ZipFile(io.BytesIO(package)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in archive.infolist():
            data = archive.read(item)
            if re.fullmatch(r"ppt/charts/chart\d+\.xml", item.filename):
                data, n = re.subn(rb'(<c:axId val="\d+"/>)(</c:(?:barChart|lineChart|areaChart)>)',
                                  rb'\1<c:axId val="99999"/>\2', data)
                assert n == 1
            z.writestr(item, data)
    deck = rpptx.Presentation.from_bytes(out.getvalue())
    shape = next(s for s in deck.slides[0].shapes if s.has_chart)
    assert tuple(shape.chart.categories) == ("x", "y")
    assert not [str(d) for d in deck.render_slide_to_png_with_diagnostics(0, 30)[1] if "chart" in str(d)]


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
@pytest.mark.parametrize("fmt", ["General", "0", "0.00", "0.0%", "#,##0", "#,##0.0", "$#,##0", '"EUR "#,##0', "# ##0"])
def test_chart_number_formats_with_separators_currency_or_literals_draw_the_chart(fmt):
    """A chart number format with a thousands separator, a currency sign or a quoted literal draws the chart. Those
    used to turn the chart into an "Unsupported chart" placeholder ("ChartML c:numFmt/@formatCode projection has
    invalid value", tensorbee/rdocx#436)."""
    prs = chart_deck(rpptx.XL_CHART_TYPE.COLUMN_CLUSTERED,
                     category_data(["Q1", "Q2"], "Revenue", (1820, 1910), number_format=fmt))
    diagnostics = [str(d) for d in prs.render_slide_to_png_with_diagnostics(0, 30)[1]]
    assert not [d for d in diagnostics if "formatCode" in d or "unsupported chart" in d]


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
@pytest.mark.parametrize("export", ["export_markdown", "export_html"])
def test_markdown_and_html_exports_hold_a_chart_as_a_table(export):
    """The Markdown and HTML exports write a chart as a table of its data: the category, the series name and the
    value are in the output. The chart used to be left out, with the diagnostic "unsupported chart, SmartArt or
    embedded-object projection omitted" (tensorbee/rdocx#436)."""
    prs = chart_deck(rpptx.XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(["Q1", "Q2"], "Zeta region", (1150, 1190)),
                     layout=5)
    prs.slides[0].shapes.title.text = "Revenue"
    text = getattr(prs, export)().text
    assert [s for s in ("Q1", "Zeta region", "1150") if s not in text] == []


def red_boxes(png):
    """Number of separate red areas along the row that crosses the most."""
    best = 0
    for row in decode_png(png):
        red = [r > 150 and g < 80 and b < 80 for r, g, b in row]
        best = max(best, sum(1 for x in range(len(red)) if red[x] and (x == 0 or not red[x - 1])))
    return best


def red_box(shapes, x):
    shape = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1), Inches(1.5), Inches(1.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = "C81E1E"
    return shape


def hide(element):
    """Set hidden="1" on the element's own p:cNvPr (the first one: the shape's or the group's)."""
    element.replace_xml(re.sub(r"<p:cNvPr ", '<p:cNvPr hidden="1" ', element.xml, count=1))


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_render_leaves_out_hidden_shapes_and_groups():
    """Shapes and groups with hidden="1", on a slide or its master, are not drawn: of four red boxes, three hidden,
    the render shows the visible one. All four used to be drawn (tensorbee/rdocx#436)."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    hide(red_box(slide.shapes, 0.5).element)
    group = slide.shapes.add_group_shape()
    red_box(group.shapes, 4)
    hide(slide.shapes[1].element)
    red_box(slide.shapes, 7.5)
    hide(red_box(prs.slide_masters[0].shapes, 10.5).element)
    assert red_boxes(prs.render_slide_to_png(0, 20)) == 1


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_pdf_gives_each_word_its_own_actual_text(tmp_path):
    """In the PDF, no /ActualText span holds a run's whole text or nothing, and a text extractor sees each word on its
    own box. The span of the first word used to carry the whole run and the others an empty /ActualText, so
    pdftotext reported "Alpha beta gamma" as one word (tensorbee/rdocx#436)."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = "Alpha beta gamma"
    pdf = prs.to_pdf()
    content = b""
    for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", pdf, re.S):
        try:
            content += zlib.decompress(m.group(1))
        except zlib.error:
            pass
    spans = [s.decode() for s in re.findall(rb"/ActualText \((.*?)\)", content)]
    assert "Alpha beta gamma" not in spans and "" not in spans
    pdftotext = poppler("pdftotext")
    if pdftotext:
        (tmp_path / "words.pdf").write_bytes(pdf)
        bbox = run([pdftotext, "-bbox", tmp_path / "words.pdf", "-"], check=True).stdout
        assert re.findall(r"<word [^>]*>([^<]*)</word>", bbox) == ["Alpha", "beta", "gamma"]


# carried by tensorbee/rdocx#442 (#436), delete once upstream main has it
def test_two_content_layout_draws_both_columns_level_and_black():
    """On the default "Two Content" layout, the right column's text starts at the left column's height and is black.
    It used to start about 2.1 in lower and grey, inheriting the master's date placeholder of the same idx
    (tensorbee/rdocx#436)."""
    dpi = 50
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(next(layout for layout in prs.slide_layouts if layout.name == "Two Content"))
    slide.shapes.title.text = "Two columns"
    slide.placeholders[1].text = "Left column text"
    slide.placeholders[2].text = "Right column text"
    rows = [[min(p) for p in row] for row in decode_png(prs.render_slide_to_png(0, dpi))]
    width, start = len(rows[0]), int(1.6 * dpi)  # below the title
    halves = []
    for xs in (range(0, width // 2), range(width // 2, width)):
        ink = [(y, min(rows[y][x] for x in xs)) for y in range(start, len(rows))]
        halves.append((next(y for y, v in ink if v < 200) / dpi, min(v for _, v in ink)))
    (left_top, _), (right_top, right_dark) = halves
    assert right_top == pytest.approx(left_top, abs=0.1)
    assert right_dark <= 40
