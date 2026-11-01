# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
DANA, SAM = "{6F1B3A52-0000-4000-8000-000000000001}", "{6F1B3A52-0000-4000-8000-000000000002}"
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)


def add_rel(rels_path, rel_type, target):
    tree = etree.parse(str(rels_path))
    rid = f"rId{max(int(r.get('Id')[3:]) for r in tree.getroot() if r.get('Id', '').startswith('rId')) + 1}"
    etree.SubElement(tree.getroot(), f"{{{NS['r']}}}Relationship", Id=rid, Type=rel_type, Target=target)
    tree.write(str(rels_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    return rid


pres = etree.parse(str(u / "ppt/presentation.xml"))
sld = pres.findall(".//p:sldId", NS)[2]  # slide 3
prels = etree.parse(str(u / "ppt/_rels/presentation.xml.rels"))
slide = u / "ppt" / prels.find(f"r:Relationship[@Id='{sld.get(RID)}']", NS).get("Target")
stree = etree.parse(str(slide))
key = next(c for c in stree.iterfind(".//p:cNvPr", NS) if c.get("name") == "Key figure")
(u / "ppt/authors.xml").write_text(
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p188:authorLst xmlns:p188="{P188}">'
    f'<p188:author id="{DANA}" name="Dana Reviewer" initials="DR" userId="Dana Reviewer" providerId="None"/>'
    f'<p188:author id="{SAM}" name="Sam Author" initials="SA" userId="Sam Author" providerId="None"/>'
    f'</p188:authorLst>')
(u / "ppt/comments").mkdir(exist_ok=True)
name = f"modernComment_{int(sld.get('id')):X}.xml"
(u / "ppt/comments" / name).write_text(
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p188:cmLst xmlns:p188="{P188}" xmlns:a="{NS["a"]}">'
    f'<p188:cm id="{{6F1B3A52-0000-4000-8000-0000000000A1}}" authorId="{DANA}" created="2026-10-01T12:00:00Z">'
    f'<ac:deMkLst xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command" '
    f'xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command">'
    f'<pc:docMk/><pc:sldMk cId="0" sldId="{sld.get("id")}"/><ac:spMk id="{key.get("id")}"/></ac:deMkLst>'
    f'<p188:replyLst><p188:reply id="{{6F1B3A52-0000-4000-8000-0000000000A2}}" authorId="{SAM}" '
    f'created="2026-10-01T12:05:00Z"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
    f'<a:t>Updated with Q3 numbers.</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>'
    f'<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Please check this figure.</a:t>'
    f'</a:r></a:p></p188:txBody></p188:cm></p188:cmLst>')
rid = add_rel(slide.parent / "_rels" / (slide.name + ".rels"),
              "http://schemas.microsoft.com/office/2018/10/relationships/comments", f"../comments/{name}")
ext_lst = stree.find("p:extLst", NS)  # PowerPoint reads a slide's modern comments through this extension
if ext_lst is None:
    ext_lst = etree.SubElement(stree.getroot(), f"{{{NS['p']}}}extLst")
ext_lst.append(etree.fromstring(
    f'<p:ext xmlns:p="{NS["p"]}" uri="{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}"><p188:commentRel '
    f'xmlns:p188="{P188}" xmlns:r="{RID[1:-3]}" r:id="{rid}"/></p:ext>'))
stree.write(str(slide), xml_declaration=True, encoding="UTF-8", standalone=True)
add_rel(u / "ppt/_rels/presentation.xml.rels", "http://schemas.microsoft.com/office/2018/10/relationships/authors",
        "authors.xml")
ct = etree.parse(str(u / "[Content_Types].xml"))
for part, kind in ((f"/ppt/comments/{name}", "comments"), ("/ppt/authors.xml", "authors")):
    etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override", PartName=part,
                     ContentType=f"application/vnd.ms-powerpoint.{kind}+xml")
ct.write(str(u / "[Content_Types].xml"), xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
