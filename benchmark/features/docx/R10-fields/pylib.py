# bench: {"tools": ["python-docx", "lxml (field results, w:updateFields)"], "processes": 1, "raw_xml": true, "unsupported": "python-docx has no layout engine: PAGE and NUMPAGES cannot be computed. The fallback, run here, writes today's date into the DATE field and sets w:updateFields so that Word refreshes every field when the file is opened (after a prompt).", "run": true, "count_loc": true}
import datetime
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

doc = docx.Document(sys.argv[1])
runs = list(doc.element.body.iter(qn("w:r")))
for i, r in enumerate(runs):
    instr = r.find(qn("w:instrText"))
    if instr is not None and instr.text.strip().startswith("DATE"):
        result = next(x for x in runs[i:] if x.find(qn("w:t")) is not None)  # first run after w:separate
        result.find(qn("w:t")).text = datetime.date.today().isoformat()
update = OxmlElement("w:updateFields")
update.set(qn("w:val"), "true")
doc.settings.element.find(qn("w:compat")).addprevious(update)
doc.save(sys.argv[2])
