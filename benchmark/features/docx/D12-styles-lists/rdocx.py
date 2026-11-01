import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
note = doc.add_style("Body Note", based_on="Normal", font_name="Arial", font_size=rdocx.Pt(10),
                     space_before=rdocx.Pt(6), space_after=rdocx.Pt(12))
for p in doc.paragraphs:
    if p.text.startswith("Note:"):
        p.style = note.style_id
k = next(i for i, p in enumerate(doc.paragraphs) if p.text == "Open the bypass.")
num_id = doc.restart_numbering(doc.paragraphs[k], start=1)
for i in (k + 1, k + 2):
    doc.paragraphs[i].numbering = (num_id, 0)
docx_ops.save_atomic(doc, out, src)
