import json
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.table import Table
from lxml import etree

src, out = sys.argv[1:3]
doc = Document(src)
result = {"body": [], "headers": [], "footers": [], "footnotes": [], "text_boxes": [], "comments": []}
for block in doc.iter_inner_content():
    if isinstance(block, Table):
        for row in block.rows:
            for cell in row.cells:
                result["body"] += [p.text for p in cell.paragraphs]
    else:
        result["body"].append(block.text)
for i, section in enumerate(doc.sections):
    for kind in ("header", "footer"):
        story = getattr(section, kind)
        if i == 0 or not story.is_linked_to_previous:
            result[kind + "s"] += [p.text for p in story.paragraphs]
result["comments"] = [c.text for c in doc.comments]
for rel in doc.part.rels.values():
    if rel.reltype == RT.FOOTNOTES:
        notes = etree.fromstring(rel.target_part.blob)
        for note in notes.iter(qn("w:footnote")):
            if note.get(qn("w:type")) in (None, "normal"):
                for p in note.iter(qn("w:p")):
                    result["footnotes"].append("".join(t.text or "" for t in p.iter(qn("w:t"))))
fallback = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
for box in doc.element.body.iter(qn("w:txbxContent")):
    if any(a.tag == fallback for a in box.iterancestors()):
        continue
    for p in box.iter(qn("w:p")):
        result["text_boxes"].append("".join(t.text or "" for t in p.iter(qn("w:t"))))
with open(out, "w") as f:
    json.dump(result, f, indent=1)
