# bench: {"tools": ["python-docx", "lxml (styles, numbering, body copy)"], "processes": 1, "raw_xml": true}
import copy
import io
import sys

import docx
from lxml import etree

# python-docx cannot merge documents: copy B's styles (renamed when A defines the same id differently, B's
# document defaults folded into its Normal), its numbering under new ids, its pictures and its body into A
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
a, b = docx.Document(sys.argv[1]), docx.Document(sys.argv[2])
xml = lambda el: etree.tostring(el)  # noqa: E731
doc_b = b.element.body

# styles: fold B's run defaults into its Normal, then rename every B style that would look different in A
sty_a, sty_b = a.styles.element, b.styles.element
dflt = lambda root: root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")  # noqa: E731
normal_b = next(s for s in sty_b.iter(W + "style") if s.get(W + "default") == "1" and s.get(W + "type") == "paragraph")
if xml(dflt(sty_a)) != xml(dflt(sty_b)):
    rpr = normal_b.find(W + "rPr")
    if rpr is None:
        rpr = etree.SubElement(normal_b, W + "rPr")
    for child in reversed(list(dflt(sty_b))):
        if rpr.find(child.tag) is None:
            rpr.insert(0, copy.deepcopy(child))
a_styles = {s.get(W + "styleId"): s for s in sty_a.iter(W + "style")}
b_styles = {s.get(W + "styleId"): s for s in sty_b.iter(W + "style")}
rename = {}
changed = True
while changed:  # a style based on a renamed style is renamed too
    changed = False
    for sid, s in b_styles.items():
        based = s.find(W + "basedOn")
        if sid not in rename and sid in a_styles and (xml(s) != xml(a_styles[sid]) or (
                based is not None and based.get(W + "val") in rename)):
            rename[sid] = "B" + sid
            changed = True
for sid, s in b_styles.items():
    if sid in a_styles and sid not in rename:
        continue
    s = copy.deepcopy(s)
    s.set(W + "styleId", rename.get(sid, sid))
    s.attrib.pop(W + "default", None)
    if sid in rename:
        s.find(W + "name").set(W + "val", "B " + s.find(W + "name").get(W + "val"))
    for ref in s.iter(W + "basedOn", W + "next", W + "link"):
        ref.set(W + "val", rename.get(ref.get(W + "val"), ref.get(W + "val")))
    sty_a.append(s)
for p in doc_b.iter(W + "p"):  # B's paragraphs without a style used B's Normal
    if p.find(f"{W}pPr/{W}pStyle") is None and normal_b.get(W + "styleId") in rename:
        ppr = p.find(W + "pPr")
        if ppr is None:
            ppr = etree.Element(W + "pPr")
            p.insert(0, ppr)
        ppr.insert(0, etree.Element(W + "pStyle", {W + "val": rename[normal_b.get(W + "styleId")]}))
for ref in doc_b.iter(W + "pStyle", W + "rStyle", W + "tblStyle"):
    ref.set(W + "val", rename.get(ref.get(W + "val"), ref.get(W + "val")))

# numbering: B's abstract and concrete lists under ids above A's
num_a, num_b = a.part.numbering_part.element, b.part.numbering_part.element
abs_off = 1 + max(int(x.get(W + "abstractNumId")) for x in num_a.iter(W + "abstractNum"))
num_off = max(int(x.get(W + "numId")) for x in num_a.iter(W + "num"))
first_num = num_a.find(W + "num")
for x in num_b.findall(W + "abstractNum"):
    x = copy.deepcopy(x)
    x.set(W + "abstractNumId", str(int(x.get(W + "abstractNumId")) + abs_off))
    first_num.addprevious(x)
for x in num_b.findall(W + "num"):
    x = copy.deepcopy(x)
    x.set(W + "numId", str(int(x.get(W + "numId")) + num_off))
    ref = x.find(W + "abstractNumId")
    ref.set(W + "val", str(int(ref.get(W + "val")) + abs_off))
    num_a.append(x)
for ref in list(doc_b.iter(W + "numId")) + [n for s in sty_a.iter(W + "style") for n in s.iter(W + "numId")
                                           if s.get(W + "styleId", "").startswith("B")]:
    if ref.get(W + "val") != "0":
        ref.set(W + "val", str(int(ref.get(W + "val")) + num_off))

# pictures: python-docx adds B's image bytes to A as a new image part
for blip in doc_b.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
    data = b.part.related_parts[blip.get(R + "embed")].blob
    blip.set(R + "embed", a.part.get_or_add_image(io.BytesIO(data))[0])

# body: B's blocks before A's final section properties
sect = a.element.body.find(W + "sectPr")
for child in list(doc_b):
    if child.tag != W + "sectPr":
        sect.addprevious(child)
a.save(sys.argv[3])
