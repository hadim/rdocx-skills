import sys

import docx_ops
import rdocx

png, out = sys.argv[1:3]
doc = rdocx.Document()
doc.update_section(0, page_width=rdocx.Mm(210), page_height=rdocx.Mm(297))
doc.add_heading("Quarterly Report", 0)
doc.add_heading("Introduction", 1)
p = doc.add_paragraph("This report has ")
p.add_run("bold findings").font.bold = True
p.add_run(" and ")
p.add_run("italic remarks").font.italic = True
p.add_run(".")
doc.add_heading("Scope", 2)
for item in ("Alpha", "Beta", "Gamma"):
    doc.add_bullet_list_item(item)
doc.add_heading("Details", 3)
for item in ("First", "Second", "Third"):
    doc.add_numbered_list_item(item)
doc.add_heading("Results", 1)
rows = [("Region", "Q1", "Q2", "Q3"), ("North", "10", "12", "14"), ("South", "8", "9", "11"),
        ("East", "7", "7", "9"), ("West", "12", "13", "15")]
table = doc.add_table(5, 4, style="Table Grid")
for r, values in enumerate(rows):
    for c, value in enumerate(values):
        table.cell(r, c).text = value
for c in range(4):
    table.cell(0, c).shading = "D9E2F3"
doc.add_picture(png, "picture.png", rdocx.Cm(8), rdocx.Cm(6))
footer = doc.section(0).footer
footer.add_paragraph("Page ").runs[0].add_field("PAGE", "1")
p = doc.section(0).footer.paragraphs[0]
p.add_run(" of ")
p.runs[-1].add_field("NUMPAGES", "1")
doc.insert_toc(1, max_level=3)
doc.rebuild_toc()
docx_ops.save_atomic(doc, out)
