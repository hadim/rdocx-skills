"""D04 inputs: a quotation template (placeholders split across runs, an item row to repeat, a total row) and
the data to fill it with (6 items)."""
import json
import sys
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import add_runs, declare_rsids, new_document, para  # noqa: E402

doc = new_document("Quotation template")
doc.add_heading("Quotation", 0)
para(doc, [("Client: ", "b"), ("{{", ""), ("client", ""), ("}}", "")])
para(doc, [("Date: ", "b"), ("{{date}}", "i")])
table = doc.add_table(rows=3, cols=3, style="Table Grid")
rows = [[("Item", "b")], [("Qty", "b")], [("Price", "b")]], \
       [[("{{name}}", "")], [("{{qty}}", "")], [("{{pri", ""), ("ce}}", "")]], \
       [[("Total", "b")], [], [("{{total}}", "b")]]
for r, cells in enumerate(rows):
    for c, parts in enumerate(cells):
        p = table.cell(r, c).paragraphs[0]
        for run in add_runs(p, parts):
            run.font.size = Pt(10)
        if c > 0:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for cell in table.rows[0].cells:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "E7E6E6")
    cell._tc.get_or_add_tcPr().append(shd)
para(doc, [("Thank you for your business, ", ""), ("{{client}}", ""), (".", "")])
declare_rsids(doc)
out = Path(sys.argv[1])
doc.save(out / "template.docx")
items = [("Steel bolts M8", 120, "0.35"), ("Nuts & washers kit", 40, "4.10"), ("Hinge, stainless", 16, "7.80"),
         ("Wood screws 4x40", 500, "0.04"), ("Wall plugs 6 mm", 300, "0.03"), ("Cable ties 200 mm", 250, "0.05")]
data = {"client": "Acme Corp", "date": "31 October 2026",
        "items": [{"name": n, "qty": str(q), "price": p} for n, q, p in items],
        "total": "471.30"}
(out / "data.json").write_text(json.dumps(data, indent=1))
