# bench: tools=python-pptx (OPC parts), lxml; processes=1; raw_xml=yes
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
a, b = Presentation(sys.argv[1]), Presentation(sys.argv[2])
picked = [b.slides[1], b.slides[2]]
# python-pptx has no import: relate A to B's slide parts, which bring their layout, master, theme, chart, image
for s in picked:  # notes stay behind (a second notes master is not allowed)
    for rid, rel in list(s.part.rels.items()):
        if rel.reltype == RT.NOTES_SLIDE:
            s.part.drop_rel(rid)
master = picked[0].slide_layout.slide_master
used = {s.slide_layout.part for s in picked}
lst = master.element.find(f"{{{P}}}sldLayoutIdLst")
for e in list(lst):  # the master keeps only the layouts the imported slides use
    if master.part.related_part(e.get(RID)) not in used:
        lst.remove(e)
        master.part.drop_rel(e.get(RID))
names = {str(p.partname) for p in a.part.package.iter_parts()}
todo, seen = [s.part for s in picked], set()
while todo:  # B's part names would collide with A's: give each one a free name
    part = todo.pop()
    if id(part) in seen:
        continue
    seen.add(id(part))
    stem, ext = re.match(r"(.*?)\d*(\.\w+)$", str(part.partname)).groups()
    n = next(n for n in range(1, 10000) if f"{stem}{n}{ext}" not in names)
    part.partname = PackURI(f"{stem}{n}{ext}")
    names.add(str(part.partname))
    todo += [rel.target_part for rel in part.rels.values() if not rel.is_external]
pres = a.part._element
ids = [int(m.get("id")) for m in pres.iter(f"{{{P}}}sldMasterId")]
ids += [int(e.get("id")) for m in a.slide_masters for e in m.element.iter(f"{{{P}}}sldLayoutId")]
next_id = max(ids) + 1
for e in lst:
    next_id += 1
    e.set("id", str(next_id))
rid = a.part.relate_to(master.part, RT.SLIDE_MASTER)
etree.SubElement(pres.find(f"{{{P}}}sldMasterIdLst"), f"{{{P}}}sldMasterId", id=str(max(ids) + 1), attrib={RID: rid})
sld_ids = a.slides._sldIdLst
for k, s in enumerate(picked):
    entry = sld_ids.add_sldId(a.part.relate_to(s.part, RT.SLIDE))
    sld_ids.remove(entry)
    sld_ids.insert(1 + k, entry)  # at position 2
a.save(sys.argv[3])
