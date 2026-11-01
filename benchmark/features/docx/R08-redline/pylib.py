# bench: {"tools": ["python-docx", "lxml + difflib (redline by hand)"], "processes": 1, "raw_xml": true}
import copy
import difflib
import re
import sys

import docx
from docx.oxml import OxmlElement
from lxml import etree

# python-docx has no tracked changes: align paragraphs with difflib, word-diff each changed one, write w:del / w:ins
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XS = "{http://www.w3.org/XML/1998/namespace}space"
ATTRS = {W + "author": "Dana Reviewer", W + "date": "2026-10-30T12:00:00Z"}
ids = iter(range(1000, 100000))
d1, d2 = docx.Document(sys.argv[1]), docx.Document(sys.argv[2])
p1, p2 = list(d1.element.body.iter(W + "p")), list(d2.element.body.iter(W + "p"))
text = lambda p: "".join(t.text or "" for t in p.iter(W + "t"))  # noqa: E731


def rev(tag, *children):
    el = OxmlElement("w:" + tag, {W + "id": str(next(ids)), **ATTRS})
    el.extend(children)
    return el


def run(rpr, s, deleted=False):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, W + ("delText" if deleted else "t"))
    t.text, _ = s, t.set(XS, "preserve")
    return r


def mark(p, tag):
    """Track the paragraph mark (w:pPr/w:rPr/w:ins or w:del)."""
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    rpr = ppr.find(W + "rPr")
    if rpr is None:
        rpr = etree.SubElement(ppr, W + "rPr")
    rpr.insert(0, rev(tag))


def delete_paragraph(p):
    for r in p.findall(W + "r"):
        for t in r.iter(W + "t"):
            t.tag = W + "delText"
        r.addprevious(rev("del"))
        r.getprevious().append(r)
    mark(p, "del")


def insert_after(anchor, src):
    p = copy.deepcopy(src)
    for r in p.findall(W + "r"):
        r.addprevious(rev("ins"))
        r.getprevious().append(r)
    mark(p, "ins")
    anchor.addnext(p)
    return p


def redline(p, new):
    """Word diff of one paragraph: its runs replaced by equal / deleted / inserted runs."""
    first = p.find(W + "r")
    rpr = first.find(W + "rPr") if first is not None else None
    a, b = re.findall(r"\S+|\s+", text(p)), re.findall(r"\S+|\s+", new)
    for child in [c for c in p if c.tag in (W + "r", W + "proofErr")]:
        p.remove(child)
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op == "equal":
            p.append(run(rpr, "".join(a[i1:i2])))
        if op in ("delete", "replace"):
            p.append(rev("del", run(rpr, "".join(a[i1:i2]), deleted=True)))
        if op in ("insert", "replace"):
            p.append(rev("ins", run(rpr, "".join(b[j1:j2]))))


last = None
sm = difflib.SequenceMatcher(None, [text(p) for p in p1], [text(p) for p in p2], autojunk=False)
for op, i1, i2, j1, j2 in sm.get_opcodes():
    if op == "equal":
        last = p1[i2 - 1]
        continue
    pairs = list(zip(range(i1, i2), range(j1, j2))) if i2 - i1 == j2 - j1 else []
    if pairs and all(difflib.SequenceMatcher(None, text(p1[i]), text(p2[j])).ratio() >= 0.5 for i, j in pairs):
        for i, j in pairs:
            redline(p1[i], text(p2[j]))
        last = p1[i2 - 1]
        continue
    for i in range(i1, i2):
        delete_paragraph(p1[i])
        last = p1[i]
    for j in range(j1, j2):
        last = insert_after(last, p2[j])
d1.save(sys.argv[3])
