import os
import re
import sys

root = sys.argv[1]


def read(name):
    return open(os.path.join(root, "word", name), encoding="utf-8").read()


def write(name, xml):
    open(os.path.join(root, "word", name), "w", encoding="utf-8").write(xml)


style = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="BodyNote"><w:name w:val="Body Note"/>'
         '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="120" w:after="240"/></w:pPr>'
         '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/>'
         '</w:rPr></w:style>')
write("styles.xml", read("styles.xml").replace("</w:styles>", style + "</w:styles>"))

numbering = read("numbering.xml")
new_id = max(int(i) for i in re.findall(r'<w:num w:numId="(\d+)"', numbering)) + 1
abstract = re.search(r'<w:num w:numId="10"><w:abstractNumId w:val="(\d+)"/>', numbering).group(1)
write("numbering.xml", numbering.replace("</w:numbering>", (
    f'<w:num w:numId="{new_id}"><w:abstractNumId w:val="{abstract}"/><w:lvlOverride w:ilvl="0">'
    '<w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>')))

doc = read("document.xml")
doc = re.sub(r"(<w:p(?: [^>]*)?>)(?=(?:(?!</w:p>).)*?<w:t>Note:</w:t>)",
             r'\1<w:pPr><w:pStyle w:val="BodyNote"/></w:pPr>', doc, flags=re.S)
start = doc.index("Then repeat for the second circuit:")
end = doc.index("Note:", start)
doc = doc[:start] + doc[start:end].replace('<w:numId w:val="10"/>', f'<w:numId w:val="{new_id}"/>') + doc[end:]
write("document.xml", doc)
