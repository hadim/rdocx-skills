# bench: {"tools": ["rdocx Python"], "processes": 1, "raw_xml": false}
import json
import sys

import rdocx

doc = rdocx.Document.open(sys.argv[1])
rows = [{"id": c.id, "author": c.author, "date": c.date, "text": c.text, "anchor": c.anchor_text,
         "parent": c.parent_id} for c in doc.comments]
open(sys.argv[2], "w").write(json.dumps({"comments": rows}, indent=1))
