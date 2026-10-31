# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
import copy
import re
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
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
NOTES = {2: "Spend two minutes on the agenda.",
         3: "Updated: the survey closed on 30 September.\nSource: field team, 412 answers.",
         4: "Stress the regional split.", 6: "Close with the three asks."}
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)


def load(path):
    return etree.parse(str(path), etree.XMLParser(remove_blank_text=False))


def save(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def slide_part(n):  # the n-th slide in <p:sldIdLst> order
    pres, rels = load(u / "ppt/presentation.xml"), load(u / "ppt/_rels/presentation.xml.rels")
    rid = pres.findall(".//p:sldId", NS)[n - 1].get("{%s}id" % RT.rstrip("/"))
    return u / "ppt" / rels.find(f"r:Relationship[@Id='{rid}']", NS).get("Target")


def set_body(notes_path, text):
    tree = load(notes_path)
    body = next(sp for sp in tree.iterfind(".//p:sp", NS)
                if sp.find(".//p:ph", NS) is not None and sp.find(".//p:ph", NS).get("type") == "body")
    tx = body.find("p:txBody", NS)
    paras = tx.findall("a:p", NS)
    template = copy.deepcopy(paras[0])
    for p in paras:
        tx.remove(p)
    for line in text.split("\n"):
        p = copy.deepcopy(template)
        runs = p.findall("a:r", NS)
        for r in runs[1:]:
            p.remove(r)
        if not runs:
            p.insert(0, etree.fromstring(f'<a:r xmlns:a="{NS["a"]}"><a:rPr lang="en-US"/><a:t/></a:r>'))
        p.find("a:r/a:t", NS).text = line
        tx.append(p)
    save(tree, notes_path)


existing = sorted(u.glob("ppt/notesSlides/notesSlide*.xml"))
for n, text in NOTES.items():
    slide = slide_part(n)
    srels_path = slide.parent / "_rels" / (slide.name + ".rels")
    srels = load(srels_path)
    rel = srels.find(f"r:Relationship[@Type='{RT}notesSlide']", NS)
    if rel is not None:
        set_body(slide.parent / rel.get("Target"), text)
        continue
    num = max(int(re.findall(r"\d+", p.stem)[0]) for p in u.glob("ppt/notesSlides/notesSlide*.xml")) + 1
    new = u / f"ppt/notesSlides/notesSlide{num}.xml"
    shutil.copy(existing[0], new)  # an existing notes slide as the template: same placeholders
    set_body(new, text)
    nrels = load(existing[0].parent / "_rels" / (existing[0].name + ".rels"))
    nrels.find(f"r:Relationship[@Type='{RT}slide']", NS).set("Target", f"../slides/{slide.name}")
    save(nrels, new.parent / "_rels" / (new.name + ".rels"))
    ids = [int(r.get("Id")[3:]) for r in srels.getroot() if r.get("Id", "").startswith("rId")]
    etree.SubElement(srels.getroot(), f"{{{NS['r']}}}Relationship", Id=f"rId{max(ids) + 1}",
                     Type=f"{RT}notesSlide", Target=f"../notesSlides/{new.name}")
    save(srels, srels_path)
    ct = load(u / "[Content_Types].xml")
    etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override", PartName=f"/ppt/notesSlides/{new.name}",
                     ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
    save(ct, u / "[Content_Types].xml")
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
