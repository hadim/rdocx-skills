# bench: tools=python-pptx, lxml; processes=1; raw_xml=yes
import copy
import sys

from pptx import Presentation

EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
prs = Presentation(sys.argv[1])
alpha, bravo, charlie, delta, echo, foxtrot, golf, hotel = list(prs.slides)
# 1. duplicate: python-pptx has none; a slide on the same layout, the source's shapes copied, their images related
dup = prs.slides.add_slide(charlie.slide_layout)
for ph in list(dup.placeholders):
    ph._element.getparent().remove(ph._element)
rids = {rel.rId: dup.part.relate_to(rel.target_part, rel.reltype)  # pictures: relate the copy to the same images
        for rel in charlie.part.rels.values() if rel.reltype.endswith("/image")}
for el in charlie.shapes._spTree.iterchildren():
    if el.tag.endswith(("}sp", "}pic", "}graphicFrame", "}grpSp", "}cxnSp")):
        el = copy.deepcopy(el)
        for node in el.iter():
            if node.get(EMBED) in rids:
                node.set(EMBED, rids[node.get(EMBED)])
        dup.shapes._spTree.append(el)
dup.notes_slide.notes_text_frame.text = charlie.notes_slide.notes_text_frame.text
# 2-3. order and deletion live in <p:sldIdLst>
lst = prs.slides._sldIdLst
entry = {s.rId: s for s in lst}
by_slide = {prs.part.related_part(rid): rid for rid in entry}
lst.remove(entry[by_slide[dup.part]])
lst.insert(lst.index(entry[by_slide[echo.part]]) + 1, entry[by_slide[dup.part]])
lst.remove(entry[by_slide[hotel.part]])
lst.insert(1, entry[by_slide[hotel.part]])
lst.remove(entry[by_slide[delta.part]])
prs.part.drop_rel(by_slide[delta.part])
# 4. hide
foxtrot._element.set("show", "0")
prs.save(sys.argv[2])
