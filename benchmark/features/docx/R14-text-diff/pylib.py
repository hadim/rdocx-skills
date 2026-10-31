# bench: {"tools": ["python-docx", "difflib"], "processes": 1, "raw_xml": false}
import difflib
import json
import sys

import docx
from docx.table import Table


def paragraphs(path):
    out = []
    for block in docx.Document(path).iter_inner_content():
        if isinstance(block, Table):
            out += [p.text for row in block.rows for cell in row.cells for p in cell.paragraphs]
        else:
            out.append(block.text)
    return out


a, b = paragraphs(sys.argv[1]), paragraphs(sys.argv[2])
out = {"changed": [], "added": [], "removed": []}
for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
    if op == "replace":
        out["changed"] += [list(pair) for pair in zip(a[i1:i2], b[j1:j2])]
        out["removed"] += a[i1 + (j2 - j1):i2]
        out["added"] += b[j1 + (i2 - i1):j2]
    elif op == "delete":
        out["removed"] += a[i1:i2]
    elif op == "insert":
        out["added"] += b[j1:j2]
open(sys.argv[3], "w").write(json.dumps(out, indent=1))
