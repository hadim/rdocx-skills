import json
import sys

import docx_ops
import rdocx

src, data_path, out = sys.argv[1:4]
data = json.load(open(data_path))
doc = rdocx.Document.open(src)
doc.replace_all([("{{client}}", data["client"], 2), ("{{date}}", data["date"], 1), ("{{total}}", data["total"], 1)])
items = data["items"]
for _ in items[1:]:
    doc.tables[0].clone_row(1)
for k, item in enumerate(items):
    for c, key in enumerate(("name", "qty", "price")):
        doc.tables[0].cell(1 + k, c).replace_text("{{%s}}" % key, item[key], expect=1)
docx_ops.save_atomic(doc, out, src)
