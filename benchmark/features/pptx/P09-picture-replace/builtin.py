# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
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
EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)
pres = etree.parse(str(u / "ppt/presentation.xml"))
prels = etree.parse(str(u / "ppt/_rels/presentation.xml.rels"))
slide = u / "ppt" / prels.find(f"r:Relationship[@Id='{pres.findall('.//p:sldId', NS)[2].get(RID)}']", NS).get("Target")
srels_path = slide.parent / "_rels" / (slide.name + ".rels")
tree, srels = etree.parse(str(slide)), etree.parse(str(srels_path))
pic = next(p for p in tree.iterfind(".//p:pic", NS) if p.find("p:nvPicPr/p:cNvPr", NS).get("name") == "Site photo")
blip = pic.find("p:blipFill/a:blip", NS)
old = blip.get(EMBED)
# a new media part: the old one is shared with slide 2, so it must not be overwritten
num = 1 + max(int("".join(filter(str.isdigit, m.stem)) or 0) for m in (u / "ppt/media").iterdir())
media = f"image{num}{Path(sys.argv[2]).suffix}"
shutil.copy(sys.argv[2], u / "ppt/media" / media)
rid = f"rId{1 + max(int(r.get('Id')[3:]) for r in srels.getroot() if r.get('Id').startswith('rId'))}"
etree.SubElement(srels.getroot(), f"{{{NS['r']}}}Relationship", Id=rid, Target=f"../media/{media}",
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
blip.set(EMBED, rid)
if not any(e.get(EMBED) == old or e.get(RID) == old for e in tree.iter()):
    srels.getroot().remove(srels.find(f"r:Relationship[@Id='{old}']", NS))
tree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
srels.write(str(srels_path), xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(sys.argv[3], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
