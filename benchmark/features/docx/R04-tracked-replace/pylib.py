# bench: {"tools": ["python-docx", "lxml (w:ins / w:del)"], "processes": 1, "raw_xml": true}
import copy
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OLD, NEW = "thirty", "forty-five"
doc = docx.Document(sys.argv[1])
p = next(p for p in doc.paragraphs if f" {OLD} " in p.text)
# python-docx has no tracked changes: find the run(s), cut the word out, wrap it in w:del, add a w:ins
start, pos = p.text.index(f" {OLD} ") + 1, 0
runs = []
for r in p.runs:
    a, b = pos, pos + len(r.text)
    pos = b
    if b > start and a < start + len(OLD):
        runs.append((r, a))
first, a0 = runs[0]
pre, rest = first.text[:start - a0], ""
for r, a in runs:
    rest = r.text[max(0, start + len(OLD) - a):] if a + len(r.text) > start + len(OLD) else rest
for r, _ in runs[1:]:
    r._r.getparent().remove(r._r)


def revision(tag, rid, run_el):
    el = OxmlElement(tag)
    for k, v in (("w:id", rid), ("w:author", "Dana Reviewer"), ("w:date", "2026-10-30T11:00:00Z")):
        el.set(qn(k), v)
    el.append(run_el)
    return el


def clone(text, deleted=False):
    r = copy.deepcopy(first._r)
    for t in r.findall(qn("w:t")):
        r.remove(t)
    t = OxmlElement("w:delText" if deleted else "w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


first.text = pre
anchor = first._r
for el in (revision("w:del", "9001", clone(OLD, True)), revision("w:ins", "9002", clone(NEW)), clone(rest)):
    anchor.addnext(el)
    anchor = el
doc.save(sys.argv[2])
