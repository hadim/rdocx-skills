import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
sentence = "The survey covered 120 sites."
i, start, _ = docx_ops.locate(doc, sentence)
first, last = docx_ops.isolate(doc, i, start, start + len(sentence))
doc.add_footnote(doc.paragraphs[i].runs[last - 1], "Source: field survey 2026.")
docx_ops.save_atomic(doc, out, src)
