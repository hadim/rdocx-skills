import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

src, out = sys.argv[1:3]


def add_field(paragraph, instruction):
    for kind, text in (("begin", None), (None, instruction), ("separate", None), (None, "1"), ("end", None)):
        run = paragraph.add_run()
        if kind:
            char = OxmlElement("w:fldChar")
            char.set(qn("w:fldCharType"), kind)
            run._r.append(char)
        elif text == instruction:
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = instruction
            run._r.append(instr)
        else:
            run.text = text


doc = Document(src)
second = doc.sections[1]
second.orientation = WD_ORIENT.LANDSCAPE
second.page_width, second.page_height = second.page_height, second.page_width
for side in ("top_margin", "right_margin", "bottom_margin", "left_margin"):
    setattr(second, side, Cm(2))
doc.sections[0].header.paragraphs[0].text = "Confidential"
footer = doc.sections[0].footer.paragraphs[0]
footer.add_run("Page ")
add_field(footer, "PAGE")
footer.add_run(" of ")
add_field(footer, "NUMPAGES")
doc.save(out)
