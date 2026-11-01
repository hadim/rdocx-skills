# bench: tools=zipfile, lxml, openpyxl; processes=1; raw_xml=yes
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from lxml import etree

NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
CATS = ["Q1 26", "Q2 26", "Q3 26", "Q4 26", "Q1 27"]
SERIES = {"North": [1250, 1330, 1410, 1535, 1600], "South": [990, 1005, 1000, 1070, 1110]}
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)


def rels(part):
    tree = etree.parse(str(part.parent / "_rels" / (part.name + ".rels")))
    return {r.get("Id"): (part.parent / r.get("Target")).resolve() for r in tree.getroot()}


pres = etree.parse(str(u / "ppt/presentation.xml"))
slide = rels(u / "ppt/presentation.xml")[pres.findall(".//p:sldId", NS)[2].get(RID)]
chart_path = rels(slide)[etree.parse(str(slide)).find(".//c:chart", NS).get(RID)]
chart = etree.parse(str(chart_path))


def fill_cache(cache, values, numeric):
    for el in cache.findall("c:pt", NS):
        cache.remove(el)
    cache.find("c:ptCount", NS).set("val", str(len(values)))
    for i, v in enumerate(values):
        pt = etree.SubElement(cache, f"{{{NS['c']}}}pt", idx=str(i))
        etree.SubElement(pt, f"{{{NS['c']}}}v").text = str(v) if numeric else v


n = len(CATS) + 1
for k, ser in enumerate(chart.iterfind(".//c:ser", NS)):  # columns B, C... rows 2..n, as the workbook lays them out
    col = "BCDEFG"[k]
    name = ser.find("c:tx//c:v", NS).text
    ser.find("c:cat//c:f", NS).text = f"Sheet1!$A$2:$A${n}"
    fill_cache(ser.find("c:cat//c:strCache", NS), CATS, False)
    ser.find("c:val//c:f", NS).text = f"Sheet1!${col}$2:${col}${n}"
    fill_cache(ser.find("c:val//c:numCache", NS), SERIES[name], True)
chart.write(str(chart_path), xml_declaration=True, encoding="UTF-8", standalone=True)
xlsx = next(p for p in rels(chart_path).values() if p.suffix == ".xlsx")
wb = openpyxl.load_workbook(xlsx)
ws = wb["Sheet1"]
for i, cat in enumerate(CATS, 2):
    ws.cell(i, 1, cat)
for k, name in enumerate(SERIES):
    ws.cell(1, 2 + k, name)
    for i, v in enumerate(SERIES[name], 2):
        ws.cell(i, 2 + k, v)
wb.save(xlsx)
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
