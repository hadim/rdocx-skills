"""Write the body fields' results: today for DATE, the page count and each PAGE field's page from LibreOffice's PDF."""
import datetime
import sys

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
doc_path, pages_path = sys.argv[1], sys.argv[2]
pages = [" ".join(p.split()) for p in open(pages_path).read().split("\f") if p.strip()]
tree = etree.parse(doc_path)
for p in tree.getroot().find(W + "body").iter(W + "p"):
    state, instr, before, results = None, "", "", []
    for r in p.iter(W + "r"):
        fc = r.find(W + "fldChar")
        kind = fc.get(W + "fldCharType") if fc is not None else None
        if kind == "begin":
            state, instr, results = "instr", "", []
        elif kind == "separate":
            state = "result"
        elif kind == "end":
            name = instr.split()[0] if instr.split() else ""
            value = {"DATE": datetime.date.today().isoformat(), "NUMPAGES": str(len(pages)),
                     "PAGE": next((str(i + 1) for i, pg in enumerate(pages) if before.strip() in pg), None)}.get(name)
            if value and results:
                results[0].text = value
                for t in results[1:]:
                    t.text = ""
            state = None
        elif state == "instr":
            instr += "".join(t.text or "" for t in r.iter(W + "instrText"))
        elif state == "result":
            results += list(r.iter(W + "t"))
        else:
            before += "".join(t.text or "" for t in r.iter(W + "t"))
tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
