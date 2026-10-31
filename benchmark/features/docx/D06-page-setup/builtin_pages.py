import os
import re
import sys

root = sys.argv[1]
WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def read(name):
    with open(os.path.join(root, name), encoding="utf-8") as f:
        return f.read()


def write(name, xml):
    with open(os.path.join(root, name), "w", encoding="utf-8") as f:
        f.write(xml)


def field(instr):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
write("word/header1.xml", DECL + f'<w:hdr {WNS}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>'
      '<w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>')
write("word/footer1.xml", DECL + f'<w:ftr {WNS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
      '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' + field("PAGE")
      + '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' + field("NUMPAGES") + '</w:p></w:ftr>')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
rels = read("word/_rels/document.xml.rels").replace("</Relationships>", (
    f'<Relationship Id="rIdHdr1" Type="{REL}header" Target="header1.xml"/>'
    f'<Relationship Id="rIdFtr1" Type="{REL}footer" Target="footer1.xml"/></Relationships>'))
write("word/_rels/document.xml.rels", rels)
CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."
types = read("[Content_Types].xml").replace("</Types>", (
    f'<Override PartName="/word/header1.xml" ContentType="{CT}header+xml"/>'
    f'<Override PartName="/word/footer1.xml" ContentType="{CT}footer+xml"/></Types>'))
write("[Content_Types].xml", types)

doc = read("word/document.xml")
refs = '<w:headerReference w:type="default" r:id="rIdHdr1"/><w:footerReference w:type="default" r:id="rIdFtr1"/>'
doc = re.sub(r"(<w:sectPr\b[^>]*>)", r"\1" + refs, doc, count=1)
last = doc.rindex("<w:sectPr")
tail = re.sub(r"<w:pgSz [^>]*/>", '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>', doc[last:], count=1)
for side in ("top", "right", "bottom", "left"):
    tail = re.sub(rf'(<w:pgMar [^>]*w:{side}=")\d+"', r'\g<1>1134"', tail, count=1)
write("word/document.xml", doc[:last] + tail)
