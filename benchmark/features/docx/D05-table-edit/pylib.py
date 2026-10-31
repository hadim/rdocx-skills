import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

src, out = sys.argv[1:3]
doc = Document(src)
table = doc.tables[0]
table.add_column(Cm(3))
for tr in table._tbl.tr_lst:
    tr.tc_lst[0].addnext(tr.tc_lst[-1])
grid = table._tbl.tblGrid
grid.gridCol_lst[0].addnext(grid.gridCol_lst[-1])
for r, value in enumerate(["Supplier", "Acme", "Bolt & Co", "Acme"]):
    table.cell(r, 1).text = value
table.cell(0, 1).paragraphs[0].runs[0].bold = True
table.cell(4, 0).merge(table.cell(4, 1))
nuts = next(row for row in table.rows if row.cells[0].text == "Nuts")
table._tbl.remove(nuts._tr)
shd = OxmlElement("w:shd")
shd.set(qn("w:val"), "clear")
shd.set(qn("w:color"), "auto")
shd.set(qn("w:fill"), "FFC000")
table.cell(0, 4)._tc.get_or_add_tcPr().append(shd)
doc.save(out)
