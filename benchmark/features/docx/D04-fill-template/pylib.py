import copy
import json
import sys

from docx import Document

src, data_path, out = sys.argv[1:4]
data = json.load(open(data_path))
doc = Document(src)


def replace(p, old, new):
    while old in "".join(r.text for r in p.runs):
        runs = p.runs
        start = "".join(r.text for r in runs).find(old)
        pos, first = 0, True
        for r in runs:
            a, b = pos, pos + len(r.text)
            pos = b
            if b > start and a < start + len(old):
                lo, hi = max(start, a) - a, min(start + len(old), b) - a
                r.text = r.text[:lo] + (new if first else "") + r.text[hi:]
                first = False


table = doc.tables[0]
template = table.rows[1]._tr
for item in data["items"]:
    tr = copy.deepcopy(template)
    template.addprevious(tr)
    for cell in table.rows[list(table._tbl.tr_lst).index(tr)].cells:
        for p in cell.paragraphs:
            for key in ("name", "qty", "price"):
                replace(p, "{{%s}}" % key, item[key])
table._tbl.remove(template)
paragraphs = list(doc.paragraphs) + [p for row in table.rows for cell in row.cells for p in cell.paragraphs]
for p in paragraphs:
    for key in ("client", "date", "total"):
        replace(p, "{{%s}}" % key, data[key])
doc.save(out)
