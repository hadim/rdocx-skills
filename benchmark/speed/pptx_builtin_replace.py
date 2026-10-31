"""Counted replacement in an unpacked .pptx (built-in stack, borrowed from features/pptx P03): every a:p of the
slides and notes slides, across runs (the replacement takes the run where the match starts); refuses, writing
nothing, when the count differs."""
import sys
from pathlib import Path

from lxml import etree

root, old, new, expect = Path(sys.argv[1]), sys.argv[2], sys.argv[3], int(sys.argv[4])
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
parts = sorted(root.glob("ppt/slides/slide*.xml")) + sorted(root.glob("ppt/notesSlides/notesSlide*.xml"))


def replace(p):
    runs = p.findall(A + "r")
    chars = [(k, c) for k, r in enumerate(runs) for c in (r.findtext(A + "t") or "")]
    text = "".join(c for _, c in chars)
    n, i, kept = 0, 0, []
    while i < len(chars):
        if text.startswith(old, i):
            kept += [(chars[i][0], c) for c in new]
            i, n = i + len(old), n + 1
        else:
            kept.append(chars[i])
            i += 1
    if n:
        for k, r in enumerate(runs):
            t = "".join(c for j, c in kept if j == k)
            if not t:
                p.remove(r)
                continue
            r.find(A + "t").text = t
            if t != t.strip():
                r.find(A + "t").set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return n


trees = {part: etree.parse(str(part)) for part in parts}
count = sum(replace(p) for tree in trees.values() for p in tree.iter(A + "p"))
if count != expect:
    sys.exit(f"refused: {count} occurrences of {old!r}, expected {expect}; nothing written")
for part, tree in trees.items():
    tree.write(str(part), xml_declaration=True, encoding="UTF-8", standalone=True)
