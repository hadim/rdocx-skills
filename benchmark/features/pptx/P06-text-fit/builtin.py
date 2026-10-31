# bench: tools=soffice.py (LibreOffice), pdftoppm, lxml, Pillow; processes=4; raw_xml=yes
# The skill's text-fit check is visual QA: render, then look for text beyond its box. A script cannot look, so it
# renders the same way and looks for ink in a band just outside each text frame that no other shape covers.
import json
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
DPI, EMU = 50, 914400
src, out = sys.argv[1], sys.argv[2]
tmp = Path(tempfile.mkdtemp())
subprocess.run([sys.executable, f"{os.environ['PPTX_SKILL']}/scripts/office/soffice.py", "--headless", "--convert-to",
                "pdf", "--outdir", tmp, src], check=True, capture_output=True)
subprocess.run(["pdftoppm", "-png", "-r", str(DPI), tmp / (Path(src).stem + ".pdf"), tmp / "slide"], check=True)
pngs = sorted(tmp.glob("slide-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))
z = zipfile.ZipFile(src)


def rels(part):
    name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    return {r.get("Id"): (r.get("Type").rsplit("/", 1)[1], posixpath.normpath(posixpath.join(posixpath.dirname(part),
            r.get("Target")))) for r in etree.fromstring(z.read(name))}


def shapes(part):
    tree = etree.fromstring(z.read(part)).find("p:cSld/p:spTree", NS)
    for sh in tree:
        if etree.QName(sh).localname in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            ph = sh.find("./*/p:nvPr/p:ph", NS)
            xfrm = sh.find("p:spPr/a:xfrm", NS)
            if xfrm is None:
                xfrm = sh.find("p:xfrm", NS) if sh.find("p:xfrm", NS) is not None else sh.find("p:grpSpPr/a:xfrm", NS)
            box = None if xfrm is None else [int(xfrm.find("a:off", NS).get(k)) for k in "xy"] + \
                [int(xfrm.find("a:ext", NS).get(k)) for k in ("cx", "cy")]
            yield {"name": sh.find("./*/p:cNvPr", NS).get("name"), "box": box, "text": "".join(sh.itertext()).strip(),
                   "ph": None if ph is None else (ph.get("type", "obj"), ph.get("idx", "0"))}


def inherited(ph, *parents):  # a placeholder without its own xfrm takes its layout's, then its master's
    for parent in parents:
        for sh in parent:
            if sh["ph"] and sh["box"] and (sh["ph"][1] == ph[1] != "0" or sh["ph"][0] == ph[0] or
                                           {sh["ph"][0], ph[0]} <= {"title", "ctrTitle"}):
                return sh["box"]
    return None


pres = etree.fromstring(z.read("ppt/presentation.xml"))
prels = rels("ppt/presentation.xml")
found = []
for n, (sld, png) in enumerate(zip(pres.iterfind("p:sldIdLst/p:sldId", NS), pngs), 1):
    slide = prels[sld.get(RID)][1]
    layout = next(t for k, t in rels(slide).values() if k == "slideLayout")
    master = next(t for k, t in rels(layout).values() if k == "slideMaster")
    lay, mas, own = list(shapes(layout)), list(shapes(master)), list(shapes(slide))
    for sh in own:
        if sh["box"] is None and sh["ph"]:
            sh["box"] = inherited(sh["ph"], lay, mas)
    decor = [s["box"] for s in lay + mas if s["ph"] is None and s["box"]]  # drawn on the slide too
    img = Image.open(png).convert("L")
    bg = max(range(256), key=img.histogram().__getitem__)
    px = lambda v: int(v / EMU * DPI)  # noqa: E731
    for sh in own:
        if not sh["text"] or sh["box"] is None:
            continue
        x, y, w, h = (px(v) for v in sh["box"])
        others = [[px(v) for v in o["box"]] for o in own if o is not sh and o["box"]] + \
                 [[px(v) for v in b] for b in decor]
        band, gap = int(0.4 * DPI), 2
        ink = 0
        for x0, y0, x1, y1 in ((x, y + h + gap, x + w, y + h + band), (x, y - band, x + w, y - gap),
                               (x + w + gap, y, x + w + band, y + h), (x - band, y, x - gap, y + h)):
            for yy in range(max(0, y0), min(img.height, y1)):
                for xx in range(max(0, x0), min(img.width, x1)):
                    if abs(img.getpixel((xx, yy)) - bg) > 60 and not any(
                            ox - gap <= xx <= ox + ow + gap and oy - gap <= yy <= oy + oh + gap for ox, oy, ow, oh in others):
                        ink += 1
        if ink > 3:
            found.append({"slide": n, "shape": sh["name"]})
with open(out, "w") as fh:
    json.dump({"overflowing": found}, fh, indent=1)
shutil.rmtree(tmp)
