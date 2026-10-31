import json
import sys

import docx_ops

src, out = sys.argv[1:3]
record = json.loads(docx_ops.cli("text", "--json", src).stdout)
result = {"body": [p["text"] for p in record["paragraphs"]], "headers": [], "footers": [], "footnotes": [],
          "text_boxes": [], "comments": []}
keys = {"header": "headers", "footer": "footers", "footnote": "footnotes", "text_box": "text_boxes",
        "comment": "comments"}
for story in record["stories"]:
    if story["kind"] in keys:
        result[keys[story["kind"]]] += [item["text"] for item in story["items"] if item["kind"] == "paragraph"]
with open(out, "w") as f:
    json.dump(result, f, indent=1)
