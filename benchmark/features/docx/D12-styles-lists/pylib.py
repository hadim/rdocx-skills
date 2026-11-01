import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

src, out = sys.argv[1:3]
doc = Document(src)
style = doc.styles.add_style("Body Note", WD_STYLE_TYPE.PARAGRAPH)
style.base_style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(10)
style.paragraph_format.space_before = Pt(6)
style.paragraph_format.space_after = Pt(12)
for p in doc.paragraphs:
    if p.text.startswith("Note:"):
        p.style = style
numbering = doc.part.numbering_part.element
k = next(i for i, p in enumerate(doc.paragraphs) if p.text == "Open the bypass.")
old = doc.paragraphs[k]._p.pPr.numPr.numId.val
num = numbering.add_num(numbering.num_having_numId(old).abstractNumId.val)
num.add_lvlOverride(ilvl=0).add_startOverride(1)
for p in doc.paragraphs[k:k + 3]:
    p._p.pPr.numPr.numId.val = num.numId
doc.save(out)
