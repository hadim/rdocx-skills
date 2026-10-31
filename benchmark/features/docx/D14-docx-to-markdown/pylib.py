import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table

src, out = sys.argv[1:3]
doc = Document(src)
numbering = doc.part.numbering_part.element


def list_format(p):
    num_pr = p._p.pPr.numPr if p._p.pPr is not None else None
    if num_pr is None or num_pr.numId is None:
        return None
    num = numbering.num_having_numId(num_pr.numId.val)
    abstract = next(a for a in numbering.findall(qn("w:abstractNum"))
                    if a.get(qn("w:abstractNumId")) == str(num.abstractNumId.val))
    level = num_pr.ilvl.val if num_pr.ilvl is not None else 0
    lvl = next(l for l in abstract.findall(qn("w:lvl")) if l.get(qn("w:ilvl")) == str(level))
    return lvl.find(qn("w:numFmt")).get(qn("w:val"))


def runs(p):
    text = ""
    for r in p.runs:
        t = r.text
        if t.strip() and r.bold:
            t = f"**{t}**"
        if t.strip() and r.italic:
            t = f"*{t}*"
        text += t
    return text


lines, counter = [], 0
for block in doc.iter_inner_content():
    if isinstance(block, Table):
        rows = [[runs(c.paragraphs[0]) for c in row.cells] for row in block.rows]
        lines += ["| " + " | ".join(rows[0]) + " |", "|" + " --- |" * len(rows[0])]
        lines += ["| " + " | ".join(r) + " |" for r in rows[1:]] + [""]
        continue
    style, fmt = block.style.name, list_format(block)
    if style.startswith("Heading"):
        lines += ["", "#" * int(style.split()[-1]) + " " + block.text, ""]
    elif fmt == "bullet":
        lines.append("- " + runs(block))
    elif fmt:
        counter += 1
        lines.append(f"{counter}. " + runs(block))
    elif block.text.strip():
        lines += ["", runs(block), ""]
    if not fmt:
        counter = 0
open(out, "w", encoding="utf-8").write("\n".join(lines) + "\n")
