import copy
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.text.run import Run

src, out = sys.argv[1:3]
doc = Document(src)
SENTENCE, NOTE = "The survey covered 120 sites.", "Source: field survey 2026."
NOTES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes {nsdecls("w")}>'
         '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
         '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p>'
         '</w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
         f'<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> {NOTE}</w:t></w:r></w:p></w:footnote></w:footnotes>')
part = Part(PackURI("/word/footnotes.xml"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", NOTES.encode(),
            doc.part.package)
doc.part.relate_to(part, RT.FOOTNOTES)

p = next(p for p in doc.paragraphs if SENTENCE in p.text)
end = p.text.index(SENTENCE) + len(SENTENCE)
pos = 0
for run in p.runs:
    if pos < end <= pos + len(run.text):
        text, cut = run.text, end - pos
        if cut < len(text):
            tail = copy.deepcopy(run._r)
            run._r.addnext(tail)
            Run(tail, p).text = text[cut:]
            run.text = text[:cut]
        run._r.addnext(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                                 '<w:footnoteReference w:id="1"/></w:r>'))
        break
    pos += len(run.text)
doc.save(out)
