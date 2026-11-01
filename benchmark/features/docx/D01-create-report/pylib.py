import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

png, out = sys.argv[1:3]


def add_field(paragraph, instruction, shown="1"):
    for kind, text in (("begin", None), (None, instruction), ("separate", None), (None, shown), ("end", None)):
        run = paragraph.add_run()
        if kind:
            char = OxmlElement("w:fldChar")
            char.set(qn("w:fldCharType"), kind)
            run._r.append(char)
        elif text == instruction:
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = instruction
            run._r.append(instr)
        else:
            run.text = text


doc = Document()
section = doc.sections[0]
section.page_width, section.page_height = Cm(21), Cm(29.7)
doc.add_heading("Quarterly Report", 0)
add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u', "Update the field to build the table of contents.")
doc.add_heading("Introduction", 1)
p = doc.add_paragraph("This report has ")
p.add_run("bold findings").bold = True
p.add_run(" and ")
p.add_run("italic remarks").italic = True
p.add_run(".")
doc.add_heading("Scope", 2)
for item in ("Alpha", "Beta", "Gamma"):
    doc.add_paragraph(item, style="List Bullet")
doc.add_heading("Details", 3)
for item in ("First", "Second", "Third"):
    doc.add_paragraph(item, style="List Number")
doc.add_heading("Results", 1)
rows = [("Region", "Q1", "Q2", "Q3"), ("North", "10", "12", "14"), ("South", "8", "9", "11"),
        ("East", "7", "7", "9"), ("West", "12", "13", "15")]
table = doc.add_table(rows=5, cols=4, style="Table Grid")
for r, values in enumerate(rows):
    for c, value in enumerate(values):
        table.cell(r, c).text = value
for cell in table.rows[0].cells:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "D9E2F3")
    cell._tc.get_or_add_tcPr().append(shd)
doc.add_picture(png, width=Cm(8))
footer = section.footer.paragraphs[0]
footer.add_run("Page ")
add_field(footer, "PAGE")
footer.add_run(" of ")
add_field(footer, "NUMPAGES")
update = OxmlElement("w:updateFields")
update.set(qn("w:val"), "true")
doc.settings.element.append(update)
doc.save(out)
