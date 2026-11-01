"""D05 input: a 5-row, 4-column priced table (bold header, a Total row) between two paragraphs."""
import sys
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import add_runs, declare_rsids, new_document, para  # noqa: E402

doc = new_document("Parts list")
doc.add_heading("Parts list", 1)
para(doc, [("The parts ordered this month are listed ", ""), ("below", "i"), (".", "")])
rows = [("Item", "Qty", "Price", "Amount"), ("Bolts", "10", "0.50", "5.00"), ("Nuts", "20", "0.20", "4.00"),
        ("Washers", "30", "0.10", "3.00"), ("Total", "", "", "12.00")]
table = doc.add_table(rows=5, cols=4, style="Table Grid")
widths = [Cm(6), Cm(3), Cm(3), Cm(4)]
for r, values in enumerate(rows):
    for c, value in enumerate(values):
        cell = table.cell(r, c)
        cell.width = widths[c]
        p = cell.paragraphs[0]
        if value:
            add_runs(p, [(value, "b" if r == 0 or (r == 4 and c in (0, 3)) else "")])
        if c > 0:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for col, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths):  # the grid matches the cells, as in Word
    col.set(qn("w:w"), str(w.twips))
para(doc, [("Prices are in euros.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
