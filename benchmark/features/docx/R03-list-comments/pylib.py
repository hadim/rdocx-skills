# bench: {"tools": ["python-docx", "lxml (anchored text)"], "processes": 1, "raw_xml": true, "notes": "python-docx gives author, date and text; the anchored text needs the range markers in the body XML"}
import json
import sys

import docx
from docx.oxml.ns import qn

doc = docx.Document(sys.argv[1])
# python-docx gives author, date and text; the anchored text is between the range markers in the body
anchors, open_ids = {}, []
for p in doc.element.body.iter(qn("w:p")):
    for cid in open_ids:
        if anchors[cid]:
            anchors[cid] += "\n"
    for el in p.iter(qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:t"), qn("w:tab")):
        if el.tag == qn("w:commentRangeStart"):
            open_ids.append(el.get(qn("w:id")))
            anchors[el.get(qn("w:id"))] = ""
        elif el.tag == qn("w:commentRangeEnd"):
            open_ids.remove(el.get(qn("w:id")))
        else:
            for cid in open_ids:
                anchors[cid] += el.text or "" if el.tag == qn("w:t") else "\t"
rows = [{"id": c.comment_id, "author": c.author, "date": c.timestamp.isoformat().replace("+00:00", "Z"),
         "text": c.text, "anchor": anchors.get(str(c.comment_id))} for c in doc.comments]
open(sys.argv[2], "w").write(json.dumps({"comments": rows}, indent=1))
