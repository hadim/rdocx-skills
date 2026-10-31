# bench: tools=python-pptx, lxml, Pillow (font metrics); processes=1; raw_xml=yes
# python-pptx has no layout engine: estimate each frame's height by wrapping its words with the fonts' metrics
# (Carlito and Caladea stand for Calibri and Cambria), sizes inherited through layout and master as PowerPoint does.
import json
import os
import sys

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
FILES = {"Calibri": "Carlito-Regular.ttf", "Cambria": "Caladea-Regular.ttf"}
prs = Presentation(sys.argv[1])
theme = prs.slide_master.part.part_related_by(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
tx = etree.fromstring(theme.blob)
FONTS = {"+mj-lt": tx.find(f".//{A}majorFont/{A}latin").get("typeface"),
         "+mn-lt": tx.find(f".//{A}minorFont/{A}latin").get("typeface")}


def font(name, pt):
    return ImageFont.truetype(os.path.join(os.environ["FONT_DIR"], FILES.get(FONTS.get(name, name), "Carlito-Regular.ttf")),
                              size=max(1, round(pt * 10)))  # measured at 10x, in tenths of a point


def inherited_size(shape, level):  # layout placeholder, then the master's title or body style; text boxes 18 pt
    if not shape.is_placeholder:
        return 18.0
    lvl = f"{A}lvl{level + 1}pPr"
    layout_ph = next((p for p in shape.part.slide_layout.placeholders
                      if p.placeholder_format.idx == shape.placeholder_format.idx), None)
    for el in ([layout_ph._element.find(f".//{A}lstStyle/{lvl}/{A}defRPr")] if layout_ph is not None else []):
        if el is not None and el.get("sz"):
            return int(el.get("sz")) / 100
    style = "titleStyle" if "TITLE" in str(shape.placeholder_format.type) else "bodyStyle"
    el = prs.slide_master._element.find(f"{P}txStyles/{P}{style}/{lvl}/{A}defRPr")
    return int(el.get("sz")) / 100


def height_needed(tf, shape, width_pt, wrap):
    total, widest = 0.0, 0.0
    for p in tf.paragraphs:
        size = next((r.font.size.pt for r in p.runs if r.font.size), None) or inherited_size(shape, p.level)
        face = next((r.font.name for r in p.runs if r.font.name), None) or ("+mj-lt" if shape.is_placeholder and
                                                                         "TITLE" in str(shape.placeholder_format.type)
                                                                         else "+mn-lt")
        f, lines, line = font(face, size), 1, ""
        for word in p.text.split():
            trial = (line + " " + word).strip()
            if wrap and f.getlength(trial) / 10 > width_pt and line:
                lines, line = lines + 1, word
            else:
                line = trial
            widest = max(widest, f.getlength(line) / 10)
        total += lines * size * 1.2
    return total, widest


def overflows(tf, shape, w, h):
    body = tf._txBody.find(f"{A}bodyPr")
    ins = {k: Emu(int(body.get(k, d))).pt for k, d in (("lIns", 91440), ("rIns", 91440), ("tIns", 45720),
                                                        ("bIns", 45720))}
    wrap = body.get("wrap") != "none"
    need, widest = height_needed(tf, shape, Emu(w).pt - ins["lIns"] - ins["rIns"], wrap)
    return need > Emu(h).pt - ins["tIns"] - ins["bIns"] or (not wrap and widest > Emu(w).pt - ins["lIns"] - ins["rIns"])


found = []
for n, slide in enumerate(prs.slides, 1):
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.TABLE:
            rows, cols = sh.table.rows, sh.table.columns
            bad = any(overflows(rows[r].cells[c].text_frame, sh, cols[c].width, rows[r].height)  # cell margins
                      for r in range(len(rows)) for c in range(len(cols)) if rows[r].cells[c].text)  # = insets
        elif sh.has_text_frame and sh.text_frame.text.strip():
            bad = overflows(sh.text_frame, sh, sh.width, sh.height)
        else:
            continue
        if bad:
            found.append({"slide": n, "shape": sh.name})
with open(sys.argv[2], "w") as fh:
    json.dump({"overflowing": found}, fh, indent=1)
