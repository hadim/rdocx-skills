import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
doc.tables[0].insert_column(0, before=False)
doc.tables[0].set_column_width(1, rdocx.Cm(3))
for r, value in enumerate(["Supplier", "Acme", "Bolt & Co", "Acme"]):
    doc.tables[0].cell(r, 1).text = value
doc.tables[0].cell(0, 1).paragraphs[0].runs[0].font.bold = True
doc.tables[0].set_cell_grid_span(4, 0, 2)
doc.tables[0].remove_row(2)
doc.tables[0].cell(0, 4).shading = "FFC000"
docx_ops.save_atomic(doc, out, src)
