# bench: {"tools": ["python-docx", "lxml (w:ins, w:del, w:rPrChange)"], "processes": 1, "raw_xml": true}
import json
import sys

import docx
from docx.oxml.ns import qn
from lxml import etree

doc = docx.Document(sys.argv[1])
# python-docx has no revision API: read w:ins / w:del / w:rPrChange in the body and in the notes parts
roots = [doc.element]
for part in doc.part.package.iter_parts():
    if str(part.partname) in ("/word/footnotes.xml", "/word/endnotes.xml"):
        roots.append(etree.fromstring(part.blob))
rows = []
for root in roots:
    for el in root.iter(qn("w:ins"), qn("w:del"), qn("w:rPrChange")):
        if el.getparent().tag == qn("w:rPr") and el.tag != qn("w:rPrChange"):
            continue  # a paragraph mark's revision holds no text
        if el.tag == qn("w:rPrChange"):
            kind, text = "format", "".join(t.text or "" for t in el.getparent().getparent().iter(qn("w:t")))
        else:
            kind = "insert" if el.tag == qn("w:ins") else "delete"
            text = "".join(t.text or "" for t in el.iter(qn("w:t"), qn("w:delText")))
        rows.append({"type": kind, "author": el.get(qn("w:author")), "date": el.get(qn("w:date")), "text": text})
open(sys.argv[2], "w").write(json.dumps({"revisions": rows}, indent=1))
