# bench: tools=python-pptx, lxml; processes=1; raw_xml=yes
import sys

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
prs = Presentation(sys.argv[1])
for slide in prs.slides:
    if slide.slide_layout.name == "Title Slide":
        continue
    for lph in slide.slide_layout.placeholders:
        kind = lph.placeholder_format.type
        if kind not in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
            continue
        slide.shapes.clone_placeholder(lph)
        ph = slide.placeholders[lph.placeholder_format.idx]
        if kind == PP_PLACEHOLDER.FOOTER:
            ph.text_frame.text = "Northwind Confidential"
        else:  # python-pptx has no field API: write the slidenum field
            p = ph.text_frame.paragraphs[0]._p
            p.append(etree.fromstring(f'<a:fld xmlns:a="{A}" id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" '
                                      f'type="slidenum"><a:rPr lang="en-US"/><a:t>&#8249;#&#8250;</a:t></a:fld>'))
prs.save(sys.argv[2])
