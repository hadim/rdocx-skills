"""Base .docx made with rdocx (the rdocx stack): styles, a table, a picture, a comment, a tracked change.

Usage: R_PY docx_rdocx.py OUT.docx PICTURE.png
Runs with common.env_rdocx() (docx_ops on PYTHONPATH). The tracked change is made the way the rdocx skill
documents it: build two versions, then `v1.compare(v2, author, timestamp)` (rdocx has no API that writes a
single w:ins / w:del by hand).
"""

import sys

import docx_ops
import rdocx

out, png = sys.argv[1], sys.argv[2]


def build(budget):
    doc = rdocx.Document()
    doc.core_properties.title = "Validity base"
    doc.core_properties.author = "bench"
    doc.add_style("Note box", based_on="Normal", italic=True, left_indent=rdocx.Inches(0.5))
    doc.add_heading("Quarterly report", 1)
    doc.add_paragraph("A styled note paragraph.", style="Note box")
    doc.add_paragraph("Revenue grew by 12 percent over the quarter.")
    doc.add_paragraph(f"The budget is {budget} thousand euros.")
    t = doc.add_table(3, 2, style="Table Grid")
    for r, row in enumerate([("Item", "Value"), ("North", "41"), ("South", "37")]):
        for c, text in enumerate(row):
            t.cell(r, c).text = text
    doc.add_paragraph("Figure 1: a picture.")
    doc.add_picture(png, "figure.png", rdocx.Inches(1.25), rdocx.Inches(0.9375))
    return doc


v1 = build("ten")
v2 = build("twelve")
v1.compare(v2, "Reviewer", "2026-10-01T09:00:00Z", granularity="word")
docx_ops.comment_on_text(v1, "Revenue grew", "Check this figure.", "Reviewer", "RV", date="2026-10-01T09:00:00Z")
docx_ops.save_atomic(v1, out)
