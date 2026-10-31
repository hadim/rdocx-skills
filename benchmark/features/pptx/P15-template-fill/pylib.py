# bench: tools=python-pptx, zipfile; processes=1; raw_xml=yes
import io
import sys
import zipfile

from pptx import Presentation

buf = io.BytesIO()  # python-pptx refuses a .potx: give it the presentation content type first
with zipfile.ZipFile(sys.argv[1]) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "[Content_Types].xml":
            data = data.replace(b"presentationml.template.main+xml", b"presentationml.presentation.main+xml")
        dst.writestr(item, data)
prs = Presentation(buf)
first = prs.slides._sldIdLst[0]  # the template's instruction slide: python-pptx cannot delete a slide
prs.part.drop_rel(first.rId)
prs.slides._sldIdLst.remove(first)
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
prs.save(sys.argv[3])
