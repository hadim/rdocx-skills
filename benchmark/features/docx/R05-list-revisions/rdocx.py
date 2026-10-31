# bench: {"tools": ["rdocx Python", "part_xml + ElementTree (revision text)"], "processes": 1, "raw_xml": true, "notes": "rdocx lists each revision's id, kind, author, date and story but not its text: the text comes from part_xml, parsed with ElementTree (NOTES_review.md, finding 4)"}
import json
import sys
import xml.etree.ElementTree as ET

import rdocx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
doc = rdocx.Document.open(sys.argv[1])
# rdocx gives each revision's id, kind, author, date and story, not its text: read that from the story's XML
text = {}
for part in {r.story.part_name for r in doc.revisions}:
    root = ET.fromstring(doc.part_xml(part))
    for el in root.iter():
        if el.tag in (W + "ins", W + "del"):
            text[(part, el.get(W + "id"))] = "".join(t.text or "" for t in el.iter() if t.tag in (W + "t", W + "delText"))
        elif el.tag == W + "r" and el.find(f"{W}rPr/{W}rPrChange") is not None:
            rid = el.find(f"{W}rPr/{W}rPrChange").get(W + "id")
            text[(part, rid)] = "".join(t.text or "" for t in el.iter(W + "t"))
kind = {"insertion": "insert", "deletion": "delete"}
rows = [{"type": kind.get(r.kind, "format"), "author": r.author, "date": r.timestamp,
         "text": text.get((r.story.part_name, str(r.id)), "")} for r in doc.revisions]
open(sys.argv[2], "w").write(json.dumps({"revisions": rows}, indent=1))
