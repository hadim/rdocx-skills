import glob
import json
import os
import sys

from lxml import etree

root, out = sys.argv[1:3]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def text(p):
    out = []
    for t in p.iter(W + "t"):
        a = t.getparent()
        while a is not p and a.tag not in (W + "txbxContent", FALLBACK):
            a = a.getparent()
        if a is p:
            out.append(t.text or "")
    return "".join(out)


def paragraphs(el, skip_boxes=True):
    out = []
    for child in el:
        if child.tag == FALLBACK or (skip_boxes and child.tag == W + "txbxContent"):
            continue
        if child.tag == W + "p":
            out.append(text(child))
        out += paragraphs(child, skip_boxes)
    return out


def part(name):
    return etree.parse(os.path.join(root, "word", name)).getroot()


doc = part("document.xml")
result = {"body": paragraphs(doc.find(W + "body")), "headers": [], "footers": [], "footnotes": [],
          "text_boxes": [], "comments": []}
for kind in ("header", "footer"):
    for name in sorted(glob.glob(os.path.join(root, "word", kind + "*.xml"))):
        result[kind + "s"] += paragraphs(part(os.path.basename(name)))
if os.path.exists(os.path.join(root, "word", "footnotes.xml")):
    for note in part("footnotes.xml").iter(W + "footnote"):
        if note.get(W + "type") in (None, "normal"):
            result["footnotes"] += paragraphs(note)
for box in doc.iter(W + "txbxContent"):
    if not any(a.tag == FALLBACK for a in box.iterancestors()):
        result["text_boxes"] += paragraphs(box, skip_boxes=False)
if os.path.exists(os.path.join(root, "word", "comments.xml")):
    for comment in part("comments.xml").iter(W + "comment"):
        result["comments"] += paragraphs(comment)
with open(out, "w") as f:
    json.dump(result, f, indent=1)
