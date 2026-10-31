"""D06 input: two sections (A4 portrait, 2.5 cm margins, no header or footer), a wide table in the second."""
import sys
from pathlib import Path

from docx.enum.section import WD_SECTION

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Annual figures")
doc.add_heading("Annual figures", 1)
para(doc, [("This first section is the ", ""), ("narrative", "i"), (" part of the report.", "")])
para(doc, [("The appendix that follows holds a ", ""), ("wide", "b"), (" table.", "")])
doc.add_section(WD_SECTION.NEW_PAGE)
doc.add_heading("Appendix: monthly figures", 1)
table = doc.add_table(rows=3, cols=13, style="Table Grid")
months = ["Region", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for c, m in enumerate(months):
    table.cell(0, c).text = m
for r, region in ((1, "North"), (2, "South")):
    table.cell(r, 0).text = region
    for c in range(1, 13):
        table.cell(r, c).text = str(10 * r + c)
para(doc, [("End of the appendix.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
