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
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)
pres = etree.parse(str(u / "ppt/presentation.xml"))
prels = etree.parse(str(u / "ppt/_rels/presentation.xml.rels"))
slide = u / "ppt" / prels.find(f"r:Relationship[@Id='{pres.findall('.//p:sldId', NS)[1].get(RID)}']", NS).get("Target")
tree = etree.parse(str(slide))
tbl = tree.find(".//a:tbl", NS)
row = copy.deepcopy(tbl.findall("a:tr", NS)[-1])  # the new row copies the last one's formatting
tbl.append(row)
cells = row.findall("a:tc", NS)
for tc, text in zip(cells, ["Total", "", "2,980", "+5 %"]):
    p = tc.find("a:txBody/a:p", NS)
    runs = p.findall("a:r", NS)
    for r in runs[1:]:
        p.remove(r)
    runs[0].find("a:t", NS).text = text
cells[0].set("gridSpan", "2")
cells[1].set("hMerge", "1")
tc = tbl.findall("a:tr", NS)[2].findall("a:tc", NS)[3]
tcpr = tc.find("a:tcPr", NS)
if tcpr is None:
    tcpr = etree.SubElement(tc, f"{{{NS['a']}}}tcPr")
for el in list(tcpr):
    if etree.QName(el).localname in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"):
        tcpr.remove(el)
fill = etree.fromstring(f'<a:solidFill xmlns:a="{NS["a"]}"><a:srgbClr val="FFC000"/></a:solidFill>')
borders = [el for el in tcpr if etree.QName(el).localname in ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr",
                                                                  "cell3D")]
tcpr.insert(tcpr.index(borders[-1]) + 1 if borders else 0, fill)  # schema order: borders, cell3D, then the fill
tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
