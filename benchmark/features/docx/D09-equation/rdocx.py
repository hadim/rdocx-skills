import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
anchor = next(p for p in doc.paragraphs if p.text == "The roots are given by:")
p = doc.insert_paragraph(doc.find_content_index(anchor) + 1, "")
result = p.add_equation(r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}", format="latex", display=True)
if result.diagnostics:
    sys.exit(f"equation diagnostics: {result.diagnostics}")
docx_ops.save_atomic(doc, out, src)
