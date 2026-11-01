# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
import copy
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships"}
FOOTER = "Northwind Confidential"
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)
for slide in sorted(u.glob("ppt/slides/slide*.xml")):
    rels = etree.parse(str(slide.parent / "_rels" / (slide.name + ".rels")))
    target = next(r for r in rels.getroot() if r.get("Type").endswith("/slideLayout")).get("Target")
    layout = etree.parse(str(slide.parent / target))
    if layout.getroot().get("type") == "title":  # "Don't show on title slide"
        continue
    tree = etree.parse(str(slide))
    sp_tree = tree.find(".//p:cSld/p:spTree", NS)
    next_id = max(int(e.get("id")) for e in tree.iterfind(".//p:cNvPr", NS)) + 1
    for sp in layout.iterfind(".//p:sp", NS):
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None or ph.get("type") not in ("ftr", "sldNum"):
            continue
        new = copy.deepcopy(sp)  # the layout's placeholder, inheriting its position and style
        new.find("p:nvSpPr/p:cNvPr", NS).set("id", str(next_id))
        next_id += 1
        new.find("p:spPr", NS).clear()
        lst = new.find("p:txBody/a:lstStyle", NS)
        lst.clear()
        if ph.get("type") == "ftr":
            p = new.find("p:txBody/a:p", NS)
            p.clear()
            r = etree.SubElement(p, f"{{{NS['a']}}}r")
            etree.SubElement(r, f"{{{NS['a']}}}rPr", lang="en-US")
            etree.SubElement(r, f"{{{NS['a']}}}t").text = FOOTER
        sp_tree.append(new)
    tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
