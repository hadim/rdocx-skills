import sys

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm

src, new_png, replacement_png, out = sys.argv[1:5]
doc = Document(src)
blip = doc.element.body.find(".//" + qn("a:blip"))
rid, _ = doc.part.get_or_add_image(replacement_png)
blip.set(qn("r:embed"), rid)
i = next(i for i, p in enumerate(doc.paragraphs) if p.text == "Figure placement anchor.")
doc.paragraphs[i + 1].insert_paragraph_before().add_run().add_picture(new_png, width=Cm(10))
doc.save(out)
