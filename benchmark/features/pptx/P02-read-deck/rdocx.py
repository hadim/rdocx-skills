# bench: tools=rpptx (Python), pptx_ops.walk; processes=1; raw_xml=no
import json
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
slides = []
for k, slide in enumerate(prs.slides):
    text = []
    for _, sh in pptx_ops.walk(slide.shapes):
        if sh.has_text_frame and sh.text:
            text.append(sh.text)
        elif sh.has_table:
            text += [cell.text for row in sh.table.rows for cell in row.cells]
    title = slide.shapes.title
    slides.append({"number": k + 1, "title": title.text if title else None, "text": text, "notes": slide.notes_text})
with open(sys.argv[2], "w") as f:
    json.dump({"slides": slides}, f, indent=1)
