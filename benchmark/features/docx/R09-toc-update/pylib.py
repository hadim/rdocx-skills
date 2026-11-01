# bench: {"tools": ["python-docx", "lxml (dirty flag, w:updateFields)"], "processes": 1, "raw_xml": true, "unsupported": "python-docx has no layout engine, so it cannot know the page numbers. The usual fallback, run here, marks the TOC field dirty and sets w:updateFields so that Word rebuilds the TOC when the file is opened (after a prompt); the file itself keeps the stale entries.", "run": true, "count_loc": true}
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

doc = docx.Document(sys.argv[1])
for instr in doc.element.body.iter(qn("w:instrText")):
    if instr.text.strip().startswith("TOC"):
        instr.getparent().getprevious().find(qn("w:fldChar")).set(qn("w:dirty"), "true")
update = OxmlElement("w:updateFields")
update.set(qn("w:val"), "true")
doc.settings.element.find(qn("w:compat")).addprevious(update)
doc.save(sys.argv[2])
