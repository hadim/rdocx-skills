# bench: {"tools": ["python-docx", "lxml (reject revisions)"], "processes": 1, "raw_xml": true}
import sys

import docx
from docx.oxml.ns import qn

AUTHOR = "Dana Reviewer"
doc = docx.Document(sys.argv[1])
# python-docx has no revision API: reject this author's w:ins / w:del / w:rPrChange by hand
for el in [e for e in doc.element.body.iter(qn("w:ins"), qn("w:del"), qn("w:rPrChange"), qn("w:pPrChange"))
           if e.get(qn("w:author")) == AUTHOR]:
    parent = el.getparent()
    if el.tag in (qn("w:rPrChange"), qn("w:pPrChange")):  # put the old properties back
        old = el.find(qn("w:rPr") if el.tag == qn("w:rPrChange") else qn("w:pPr"))
        for child in list(parent):
            parent.remove(child)
        for child in (list(old) if old is not None else []):
            parent.append(child)
    elif parent.tag == qn("w:rPr") and el.tag == qn("w:del"):  # a deleted paragraph mark comes back
        parent.remove(el)
    elif parent.tag == qn("w:rPr"):  # an inserted paragraph mark goes: join the paragraph to the next one
        p = parent.getparent().getparent()
        parent.remove(el)
        nxt = p.getnext()
        if nxt is not None and nxt.tag == qn("w:p"):
            ppr = nxt.find(qn("w:pPr"))
            for i, c in enumerate([c for c in p if c.tag != qn("w:pPr")]):
                nxt.insert(i + (ppr is not None), c)
            p.getparent().remove(p)
    elif el.tag == qn("w:ins"):  # inserted text goes
        parent.remove(el)
    else:  # deleted text comes back as plain text
        for t in el.iter(qn("w:delText")):
            t.tag = qn("w:t")
        i = parent.index(el)
        for c in reversed(list(el)):
            parent.insert(i, c)
        parent.remove(el)
doc.save(sys.argv[2])
