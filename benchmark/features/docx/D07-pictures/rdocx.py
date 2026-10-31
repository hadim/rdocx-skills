import sys

import docx_ops
import rdocx

src, new_png, replacement_png, out = sys.argv[1:5]
doc = rdocx.Document.open(src)
doc.replace_image(doc.pictures()[0].relationship_id, open(replacement_png, "rb").read())
anchor = next(it for it in doc.story_items if it.kind == "paragraph" and it.text == "Figure placement anchor.")
data = open(new_png, "rb").read()
w, h = docx_ops.image_size(data)
doc.add_picture(data, "new.png", rdocx.Cm(10), rdocx.Cm(10 * h / w), after=anchor)
docx_ops.save_atomic(doc, out, src)
