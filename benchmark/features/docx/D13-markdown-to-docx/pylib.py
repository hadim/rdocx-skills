import re
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

src, out = sys.argv[1:3]
doc = Document()
INLINE = re.compile(r"\*\*(.+?)\*\*|\*(.+?)\*|\[(.+?)\]\((.+?)\)|`(.+?)`")


def inline(p, text):
    pos = 0
    for m in INLINE.finditer(text):
        p.add_run(text[pos:m.start()])
        bold, italic, label, url, code = m.groups()
        if bold:
            p.add_run(bold).bold = True
        elif italic:
            p.add_run(italic).italic = True
        elif code:
            p.add_run(code).font.name = "Courier New"
        else:
            link = OxmlElement("w:hyperlink")
            link.set(qn("r:id"), doc.part.relate_to(url, RT.HYPERLINK, is_external=True))
            run = p.add_run(label)
            run.font.underline, run.font.color.rgb = True, RGBColor(0x05, 0x63, 0xC1)
            link.append(run._r)
            p._p.append(link)
        pos = m.end()
    p.add_run(text[pos:])


lines = open(src, encoding="utf-8").read().splitlines()
i = 0
while i < len(lines):
    line = lines[i]
    if line.startswith("```"):
        i += 1
        while not lines[i].startswith("```"):
            p = doc.add_paragraph(style="No Spacing")
            run = p.add_run(lines[i])
            run.font.name, run.font.size = "Courier New", Pt(9)
            i += 1
    elif m := re.match(r"(#{1,6}) (.*)", line):
        doc.add_heading(m.group(2), len(m.group(1)))
    elif m := re.match(r"[-*] (.*)", line):
        inline(doc.add_paragraph(style="List Bullet"), m.group(1))
    elif m := re.match(r"\d+\. (.*)", line):
        inline(doc.add_paragraph(style="List Number"), m.group(1))
    elif line.startswith("|"):
        rows = []
        while i < len(lines) and lines[i].startswith("|"):
            if not re.match(r"\|[\s:|-]+\|$", lines[i]):
                rows.append([c.strip() for c in lines[i].strip("|").split("|")])
            i += 1
        table = doc.add_table(rows=len(rows), cols=len(rows[0]), style="Table Grid")
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                inline(table.cell(r, c).paragraphs[0], text)
                if r == 0:
                    table.cell(r, c).paragraphs[0].runs[0].bold = True
        continue
    elif line.strip():
        inline(doc.add_paragraph(), line)
    i += 1
doc.save(out)
