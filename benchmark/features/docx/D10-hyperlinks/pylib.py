import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

src, out = sys.argv[1:3]
doc = Document(src)


def link(paragraph, text, **attrs):
    hyperlink = OxmlElement("w:hyperlink")
    for key, value in attrs.items():
        hyperlink.set(qn(key), value)
    run = paragraph.add_run(text)
    run.font.underline = True
    run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)


p = next(p for p in doc.paragraphs if p.text.startswith("Further reading:"))
heading = next(h for h in doc.paragraphs if h.text == "Appendix A")
ids = [int(b.get(qn("w:id"))) for b in doc.element.body.iter(qn("w:bookmarkStart"))]
start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
start.set(qn("w:id"), str(max(ids, default=0) + 1))
start.set(qn("w:name"), "_Appendix_A")
end.set(qn("w:id"), start.get(qn("w:id")))
heading.runs[0]._r.addprevious(start)
heading.runs[-1]._r.addnext(end)
rid = doc.part.relate_to("https://example.com/project", RT.HYPERLINK, is_external=True)
link(p, "Project site", **{"r:id": rid})
p.add_run(" and ")
link(p, "Appendix A", **{"w:anchor": "_Appendix_A"})
p.add_run(".")
doc.save(out)
