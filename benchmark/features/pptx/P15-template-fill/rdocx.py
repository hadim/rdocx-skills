# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
prs.slides.remove(prs.slides[0])  # the template's instruction slide
layout = {lay.name: lay for lay in prs.slide_layouts}
s = prs.slides.add_slide(layout["Title Slide"])
s.shapes.title.text = "Northwind annual review"
s.placeholders[1].text = "Fiscal year 2026"
s = prs.slides.add_slide(layout["Title and Content"])
s.shapes.title.text = "Highlights"
body = s.placeholders[1].text_frame
body.text = "Revenue up 12 %"
for line in ("Two new regions", "Customer churn down to 4 %"):
    body.add_paragraph().text = line
s = prs.slides.add_slide(layout["Picture with Caption"])
s.shapes.title.text = "Our new headquarters"
s.placeholders[1].insert_picture(sys.argv[2])  # cropped to fill the placeholder
s.placeholders[2].text = "Opening March 2027"
pptx_ops.save_atomic(prs, sys.argv[3], sys.argv[1])  # saved as .pptx: a presentation
