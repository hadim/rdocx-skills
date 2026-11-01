# bench: tools=add_slide.py, clean.py, zipfile, lxml; processes=3; raw_xml=yes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

SCRIPTS = Path(os.environ["PPTX_SKILL"]) / "scripts"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(u)


def load(path):
    return etree.parse(str(path))


def save(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def rels_path(part):
    return part.parent / "_rels" / (part.name + ".rels")


def slide_files():  # slide file names in <p:sldIdLst> order
    pres, prels = load(u / "ppt/presentation.xml"), load(u / "ppt/_rels/presentation.xml.rels")
    targets = {r.get("Id"): Path(r.get("Target")).name for r in prels.getroot()}
    return [targets[s.get(RID)] for s in pres.iterfind(".//p:sldId", NS)]


orig = slide_files()  # original numbering: orig[2] is slide 3
# 1. duplicate slide 3 after slide 5; add_slide.py leaves the copy without notes, so copy them too
out = subprocess.run([sys.executable, SCRIPTS / "add_slide.py", u, orig[2], "--after", orig[4]],
                     capture_output=True, text=True, check=True).stdout
copy = u / re.search(r"Created (ppt/slides/\S+\.xml)", out).group(1)
src_notes = next(r.get("Target") for r in load(rels_path(u / "ppt/slides" / orig[2])).getroot()
                 if r.get("Type") == RT + "notesSlide")
src_notes = (u / "ppt/slides" / src_notes).resolve()
n = 1 + max(int(re.findall(r"\d+", p.stem)[0]) for p in (u / "ppt/notesSlides").glob("notesSlide*.xml"))
new_notes = u / f"ppt/notesSlides/notesSlide{n}.xml"
shutil.copy(src_notes, new_notes)
nrels = load(rels_path(src_notes))
next(r for r in nrels.getroot() if r.get("Type") == RT + "slide").set("Target", f"../slides/{copy.name}")
save(nrels, rels_path(new_notes))
crels = load(rels_path(copy))
ids = [int(r.get("Id")[3:]) for r in crels.getroot() if r.get("Id").startswith("rId")]
etree.SubElement(crels.getroot(), f"{{{NS['r']}}}Relationship", Id=f"rId{max(ids) + 1}", Type=RT + "notesSlide",
                 Target=f"../notesSlides/{new_notes.name}")
save(crels, rels_path(copy))
ct = load(u / "[Content_Types].xml")
etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override", PartName=f"/ppt/notesSlides/{new_notes.name}",
                 ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
save(ct, u / "[Content_Types].xml")
# 2. move the last slide to position 2, 3. delete slide 4: edit <p:sldIdLst>
pres, prels = load(u / "ppt/presentation.xml"), load(u / "ppt/_rels/presentation.xml.rels")
by_file = {Path(r.get("Target")).name: r.get("Id") for r in prels.getroot()}
lst = pres.find(".//p:sldIdLst", NS)
entry = {s.get(RID): s for s in lst}
lst.remove(entry[by_file[orig[-1]]])
lst.insert(1, entry[by_file[orig[-1]]])
lst.remove(entry[by_file[orig[3]]])
save(pres, u / "ppt/presentation.xml")
subprocess.run([sys.executable, SCRIPTS / "clean.py", u], capture_output=True, check=True)
# 4. hide slide 6
hidden = load(u / "ppt/slides" / orig[5])
hidden.getroot().set("show", "0")
save(hidden, u / "ppt/slides" / orig[5])
with zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
