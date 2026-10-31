# bench: {"tools": ["python-docx", "lxml (accept revisions)"], "processes": 1, "raw_xml": true}
import sys

import docx
from docx.oxml.ns import qn
from lxml import etree

doc = docx.Document(sys.argv[1])


def accept(root):
    """Accept every tracked change under root (python-docx has no API for revisions)."""
    for tag in ("w:del", "w:moveFrom"):  # deleted content goes; a deleted paragraph mark joins the next one
        for el in list(root.iter(qn(tag))):
            if el.getparent().tag == qn("w:rPr") and el.getparent().getparent().tag == qn("w:pPr"):
                p = el.getparent().getparent().getparent()
                nxt = p.getnext()
                el.getparent().remove(el)
                if nxt is not None and nxt.tag == qn("w:p"):
                    content = [c for c in p if c.tag != qn("w:pPr")]
                    ppr = nxt.find(qn("w:pPr"))
                    pos = 0 if ppr is None else 1
                    for c in reversed(content):
                        nxt.insert(pos, c)
                    p.getparent().remove(p)
            elif el.getparent() is not None:
                el.getparent().remove(el)
    for tag in ("w:ins", "w:moveTo"):  # inserted content stays, unwrapped
        for el in list(root.iter(qn(tag))):
            parent = el.getparent()
            if parent.tag == qn("w:rPr"):
                parent.remove(el)
                continue
            i = parent.index(el)
            for c in reversed(list(el)):
                parent.insert(i, c)
            parent.remove(el)
    for tag in ("w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange", "w:trPrChange", "w:tcPrChange"):
        for el in list(root.iter(qn(tag))):
            el.getparent().remove(el)


accept(doc.element)
for part in doc.part.package.iter_parts():  # notes, headers and footers python-docx keeps as XML blobs
    if str(part.partname) in ("/word/footnotes.xml", "/word/endnotes.xml"):
        root = etree.fromstring(part.blob)
        accept(root)
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
doc.save(sys.argv[2])
