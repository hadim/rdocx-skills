# bench: {"tools": ["rdocx Python (layout)"], "processes": 1, "raw_xml": false}
import json
import os
import sys

import rdocx

doc = rdocx.Document.open(sys.argv[1])
first_page = {}
for f in doc.layout(font_dir=os.environ["FONT_DIR"]):
    first_page.setdefault(f.body_index, f.physical_page)
rows = [{"level": int(p.style[-1]), "text": p.text, "page": first_page[doc.find_content_index(p)]}
        for p in doc.paragraphs if (p.style or "").startswith("Heading")]
open(sys.argv[2], "w").write(json.dumps({"headings": rows}, indent=1))
