import sys

path = sys.argv[1]
xml = open(path, encoding="utf-8").read()


def r(text):
    return f"<m:r><m:t>{text}</m:t></m:r>"


radical = ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e>' + r("b") + "</m:e><m:sup>"
           + r("2") + "</m:sup></m:sSup>" + r("\u2212" + "4ac") + "</m:e></m:rad>")
equation = ('<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>'
            + r("x=") + "<m:f><m:num>" + r("\u2212b\u00b1") + radical + "</m:num><m:den>" + r("2a")
            + "</m:den></m:f></m:oMath></m:oMathPara></w:p>")
anchor = xml.index("The roots are given by:")
end = xml.index("</w:p>", anchor) + len("</w:p>")
open(path, "w", encoding="utf-8").write(xml[:end] + equation + xml[end:])
