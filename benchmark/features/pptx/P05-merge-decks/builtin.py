# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
A, B = Path(tempfile.mkdtemp()), Path(tempfile.mkdtemp())
zipfile.ZipFile(sys.argv[1]).extractall(A)
zipfile.ZipFile(sys.argv[2]).extractall(B)


def load(path):
    return etree.parse(str(path))


def save(tree, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def rels_name(part):
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def rels(root, part):  # [(rel element, absolute target part or None when external)]
    path = root / rels_name(part)
    if not path.exists():
        return []
    return [(r, None if r.get("TargetMode") == "External"
             else posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get("Target"))))
            for r in load(path).getroot()]


def kept(part, rel, target):  # notes stay behind; a master keeps only the layouts the imported slides use
    if rel.get("Type") == RT + "notesSlide":
        return False
    return not (rel.get("Type") == RT + "slideLayout" and "slideMaster" in part and target not in used_layouts)


bpres = load(B / "ppt/presentation.xml")
brels = {r.get("Id"): t for r, t in rels(B, "ppt/presentation.xml")}
picked = [brels[s.get(RID)] for s in bpres.findall(".//p:sldId", NS)[1:3]]  # B's slides 2 and 3
used_layouts = {t for s in picked for r, t in rels(B, s) if r.get("Type") == RT + "slideLayout"}
mapping, todo = {}, list(picked)
while todo:  # every part the picked slides need, under a name free in A
    part = todo.pop(0)
    if part in mapping:
        continue
    stem, ext = re.match(r"(.*?)\d*(\.\w+)$", part).groups()
    n = 1
    while (A / f"{stem}{n}{ext}").exists() or f"{stem}{n}{ext}" in mapping.values():
        n += 1
    mapping[part] = f"{stem}{n}{ext}"
    todo += [t for r, t in rels(B, part) if t and kept(part, r, t)]
bct, act = load(B / "[Content_Types].xml"), load(A / "[Content_Types].xml")
overrides = {o.get("PartName").lstrip("/"): o.get("ContentType") for o in bct.getroot() if o.get("PartName")}
defaults = {d.get("Extension") for d in act.getroot() if d.get("Extension")}
for src, dst in mapping.items():
    (A / dst).parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(B / src, A / dst)
    out = etree.Element(f"{{{NS['r']}}}Relationships", nsmap={None: NS["r"]})
    for r, t in rels(B, src):
        if t is None or kept(src, r, t):
            r = etree.SubElement(out, r.tag, dict(r.attrib))
            if t:
                r.set("Target", posixpath.relpath(mapping[t], posixpath.dirname(dst)))
    if len(out):
        save(etree.ElementTree(out), A / rels_name(dst))
    if src in overrides:
        etree.SubElement(act.getroot(), f"{{{NS['ct']}}}Override", PartName="/" + dst, ContentType=overrides[src])
    elif dst.rsplit(".", 1)[1] not in defaults:
        d = next(d for d in bct.getroot() if d.get("Extension") == dst.rsplit(".", 1)[1])
        act.getroot().insert(0, etree.fromstring(etree.tostring(d)))
        defaults.add(d.get("Extension"))
save(act, A / "[Content_Types].xml")
apres, arels_path = load(A / "ppt/presentation.xml"), A / "ppt/_rels/presentation.xml.rels"
arels = load(arels_path)
next_rid = max(int(r.get("Id")[3:]) for r in arels.getroot() if r.get("Id").startswith("rId")) + 1
ids = [int(m.get("id")) for m in apres.iterfind(".//p:sldMasterId", NS)]
for m in apres.iterfind(".//p:sldMasterId", NS):
    master = {r.get("Id"): t for r, t in rels(A, "ppt/presentation.xml")}[m.get(RID)]
    ids += [int(e.get("id")) for e in load(A / master).iterfind(".//p:sldLayoutId", NS)]
next_id = max(ids) + 1
for src, dst in mapping.items():
    if "slideMasters/" not in dst:
        continue
    master = load(A / dst)  # its layout list: only the layouts copied, with ids unique in A
    keep = {r.get("Id") for r, t in rels(A, dst)}
    master_id = next_id
    for e in master.findall(".//p:sldLayoutId", NS):
        if e.get(RID) not in keep:
            e.getparent().remove(e)
        else:
            next_id += 1
            e.set("id", str(next_id))
    save(master, A / dst)
    etree.SubElement(arels.getroot(), f"{{{NS['r']}}}Relationship", Id=f"rId{next_rid}", Type=RT + "slideMaster",
                     Target=posixpath.relpath(dst, "ppt"))
    etree.SubElement(apres.find("p:sldMasterIdLst", NS), f"{{{NS['p']}}}sldMasterId", id=str(master_id),
                     attrib={RID: f"rId{next_rid}"})
    next_rid, next_id = next_rid + 1, next_id + 1
lst = apres.find("p:sldIdLst", NS)
sld_id = max(int(s.get("id")) for s in lst) + 1
for k, src in enumerate(picked):
    etree.SubElement(arels.getroot(), f"{{{NS['r']}}}Relationship", Id=f"rId{next_rid}", Type=RT + "slide",
                     Target=posixpath.relpath(mapping[src], "ppt"))
    entry = etree.Element(f"{{{NS['p']}}}sldId", id=str(sld_id + k), attrib={RID: f"rId{next_rid}"})
    lst.insert(1 + k, entry)  # at position 2
    next_rid += 1
save(apres, A / "ppt/presentation.xml")
save(arels, arels_path)
with zipfile.ZipFile(sys.argv[3], "w", zipfile.ZIP_DEFLATED) as z:
    z.write(A / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(A.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(A).as_posix())
shutil.rmtree(A)
shutil.rmtree(B)
