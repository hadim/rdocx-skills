"""Reject one author's tracked changes in document.xml (no helper script does it): drop their insertions,
restore their deletions, undo their formatting changes, rejoin a paragraph they inserted."""
import sys

from lxml import etree

path, author = sys.argv[1], sys.argv[2]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
tree = etree.parse(path)
mine = lambda el: el.get(W + "author") == author  # noqa: E731
for el in [e for e in tree.iter(W + "ins", W + "del", W + "rPrChange", W + "pPrChange") if mine(e)]:
    parent = el.getparent()
    if el.tag in (W + "rPrChange", W + "pPrChange"):  # put the old properties back
        old = el.find(W + ("rPr" if el.tag == W + "rPrChange" else "pPr"))
        for child in list(parent):
            parent.remove(child)
        for child in (list(old) if old is not None else []):
            parent.append(child)
    elif parent.tag == W + "rPr" and el.tag == W + "del":  # a deleted paragraph mark comes back
        parent.remove(el)
    elif parent.tag == W + "rPr":  # an inserted paragraph mark goes: join the paragraph to the next one
        p = parent.getparent().getparent()
        parent.remove(el)
        nxt = p.getnext()
        if nxt is not None and nxt.tag == W + "p":
            ppr = nxt.find(W + "pPr")
            for i, c in enumerate([c for c in p if c.tag != W + "pPr"]):
                nxt.insert(i + (ppr is not None), c)
            p.getparent().remove(p)
    elif el.tag == W + "ins":  # inserted text goes
        parent.remove(el)
    else:  # deleted text comes back as plain text
        for t in el.iter(W + "delText"):
            t.tag = W + "t"
        i = parent.index(el)
        for c in reversed(list(el)):
            parent.insert(i, c)
        parent.remove(el)
tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
