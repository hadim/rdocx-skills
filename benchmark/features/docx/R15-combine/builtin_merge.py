"""Append B's body to A (no helper script merges documents): B's styles copied, renamed with a B prefix when A
defines the same id differently (B's document defaults folded into its Normal first), B's numbering copied under
new ids, B's pictures copied under new part names, then B's paragraphs."""
import copy
import shutil
import sys
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
a, b = Path(sys.argv[1]) / "word", Path(sys.argv[2]) / "word"
trees = {}
for side, folder in (("a", a), ("b", b)):
    for name in ("document.xml", "styles.xml", "numbering.xml", "_rels/document.xml.rels"):
        trees[side, name] = etree.parse(str(folder / name))
xml = lambda el: etree.tostring(el)  # noqa: E731
doc_b = trees["b", "document.xml"].getroot().find(W + "body")

# styles: fold B's run defaults into its Normal, then rename every B style that would look different in A
sty_a, sty_b = trees["a", "styles.xml"].getroot(), trees["b", "styles.xml"].getroot()
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
num_a, num_b = trees["a", "numbering.xml"].getroot(), trees["b", "numbering.xml"].getroot()
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

# pictures: B's media under new names, related from A's document part
rels_a, rels_b = trees["a", "_rels/document.xml.rels"].getroot(), trees["b", "_rels/document.xml.rels"].getroot()
new_rid = {}
for k, rel in enumerate(rels_b.findall(PR + "Relationship")):
    if rel.get("Type").endswith("/image"):
        target = "media/b_" + Path(rel.get("Target")).name
        shutil.copy(b / rel.get("Target"), a / target)
        new_rid[rel.get("Id")] = f"rIdB{k}"
        etree.SubElement(rels_a, PR + "Relationship", {"Id": f"rIdB{k}", "Type": rel.get("Type"), "Target": target})
for blip in doc_b.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
    blip.set(R + "embed", new_rid[blip.get(R + "embed")])

# body: B's blocks before A's final section properties
body_a = trees["a", "document.xml"].getroot().find(W + "body")
sect = body_a.find(W + "sectPr")
for child in doc_b:
    if child.tag != W + "sectPr":
        sect.addprevious(child)
for (side, name), tree in trees.items():
    if side == "a":
        tree.write(str(a / name), xml_declaration=True, encoding="UTF-8", standalone=True)
