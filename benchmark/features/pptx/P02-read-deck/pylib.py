# bench: tools=python-pptx; processes=1; raw_xml=no
import json
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def texts(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from texts(sh.shapes)
        elif sh.has_text_frame and sh.text_frame.text:
            yield sh.text_frame.text
        elif getattr(sh, "has_table", False):
            yield from (cell.text for row in sh.table.rows for cell in row.cells)


prs = Presentation(sys.argv[1])
slides = []
for k, slide in enumerate(prs.slides):
    title = slide.shapes.title
    notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else None
    slides.append({"number": k + 1, "title": title.text if title else None, "text": list(texts(slide.shapes)),
                   "notes": notes})
with open(sys.argv[2], "w") as f:
    json.dump({"slides": slides}, f, indent=1)
