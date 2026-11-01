import sys

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

src, out = sys.argv[1:3]
doc = Document(src)


def r(text):
    return f"<m:r><m:t>{text}</m:t></m:r>"


radical = ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e>' + r("b") + "</m:e><m:sup>"
           + r("2") + "</m:sup></m:sSup>" + r("\u2212" + "4ac") + "</m:e></m:rad>")
equation = (f'<w:p {nsdecls("w", "m")}><m:oMathPara><m:oMath>' + r("x=") + "<m:f><m:num>" + r("\u2212b\u00b1")
            + radical + "</m:num><m:den>" + r("2a") + "</m:den></m:f></m:oMath></m:oMathPara></w:p>")
anchor = next(p for p in doc.paragraphs if p.text == "The roots are given by:")
anchor._p.addnext(parse_xml(equation))
doc.save(out)
