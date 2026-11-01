import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
two = rdocx.Cm(2)
doc.update_section(1, orientation="landscape", margin_top=two, margin_right=two, margin_bottom=two, margin_left=two)
doc.section(0).header.add_paragraph("Confidential")
doc.section(0).footer.add_paragraph("Page ").runs[0].add_field("PAGE", "1")
footer = doc.section(0).footer.paragraphs[0]
footer.add_run(" of ")
footer.runs[-1].add_field("NUMPAGES", "1")
docx_ops.save_atomic(doc, out, src)
