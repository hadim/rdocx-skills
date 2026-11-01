import glob
import os
import sys

from lxml import etree

root, old, new, expected = sys.argv[1:5]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def replace(p):
    count = 0
    while True:
        ts = list(p.iter(W + "t"))
        text = "".join(t.text or "" for t in ts)
        start = text.find(old)
        if start < 0:
            return count
        pos, first = 0, True
        for t in ts:
            s = t.text or ""
            a, b = pos, pos + len(s)
            pos = b
            if b > start and a < start + len(old):
                lo, hi = max(start, a) - a, min(start + len(old), b) - a
                t.text = s[:lo] + (new if first else "") + s[hi:]
                t.set(SPACE, "preserve")
                first = False
        count += 1


parts = [os.path.join(root, "word", "document.xml")]
parts += sorted(glob.glob(os.path.join(root, "word", "header*.xml")) + glob.glob(os.path.join(root, "word", "footer*.xml")))
trees = {path: etree.parse(path) for path in parts}
total = sum(replace(p) for tree in trees.values() for p in tree.iter(W + "p"))
if total != int(expected):
    sys.exit(f"refused: expected {expected} occurrence(s) of {old!r}, found {total}; nothing written")
for path, tree in trees.items():
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
