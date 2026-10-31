# bench: {"tools": ["python-docx", "lxml (markers, commentsExtended)"], "processes": 1, "raw_xml": true}
import copy
import sys
from datetime import datetime, timezone

import docx
from docx.oxml.ns import qn
from lxml import etree

doc = docx.Document(sys.argv[1])
parent = next(c for c in doc.comments if c.author == "Alex Author")
reply = doc.comments.add_comment(text="Agreed, changed to 45 days.", author="Dana Reviewer", initials="DR")
reply._comment_elm.date = datetime(2026, 10, 30, 10, 0, tzinfo=timezone.utc)  # no public setter
pid, rid = str(parent.comment_id), str(reply.comment_id)
# python-docx threads nothing: put the reply's markers next to the parent's, as Word does
body = doc.element.body
for tag, side in (("w:commentRangeStart", "addnext"), ("w:commentRangeEnd", "addprevious")):
    mark = body.xpath(f"//{tag}[@w:id='{pid}']")[0]
    new = copy.deepcopy(mark)
    new.set(qn("w:id"), rid)
    getattr(mark, side)(new)
ref_run = body.xpath(f"//w:commentReference[@w:id='{pid}']")[0].getparent()
new_ref = copy.deepcopy(ref_run)
new_ref.xpath("w:commentReference")[0].set(qn("w:id"), rid)
ref_run.addnext(new_ref)
# the thread and its resolved state live in commentsExtended.xml, which python-docx keeps as a blob
reply_p = reply._comment_elm.xpath("w:p")[-1]
reply_p.set(qn("w14:paraId"), "7E3A1C01")
parent_para = parent._comment_elm.xpath("w:p")[-1].get(qn("w14:paraId"))
part = next(r.target_part for r in doc.part.rels.values() if r.reltype.endswith("/commentsExtended"))
root = etree.fromstring(part.blob)
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
for ex in root:
    if ex.get(f"{{{W15}}}paraId") == parent_para:
        ex.set(f"{{{W15}}}done", "1")
etree.SubElement(root, f"{{{W15}}}commentEx", {f"{{{W15}}}paraId": "7E3A1C01",
                                                f"{{{W15}}}paraIdParent": parent_para, f"{{{W15}}}done": "0"})
part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
doc.save(sys.argv[2])
