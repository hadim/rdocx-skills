import os
import re
import sys

root = sys.argv[1]
doc_path = os.path.join(root, "word", "document.xml")
rels_path = os.path.join(root, "word", "_rels", "document.xml.rels")
doc = open(doc_path, encoding="utf-8").read()
rels = open(rels_path, encoding="utf-8").read()

rels = rels.replace("</Relationships>", '<Relationship Id="rIdLink1" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/hyperlink" Target="https://example.com/project" '
                    'TargetMode="External"/></Relationships>')
bid = max([int(i) for i in re.findall(r'<w:bookmarkStart w:id="(\d+)"', doc)], default=0) + 1
heading = re.search(r'(<w:p[ >](?:(?!<w:p[ >]).)*?<w:pStyle w:val="Heading1"/></w:pPr>)(<w:r>.{0,60}?'
                    r'<w:t>Appendix A</w:t></w:r>)(</w:p>)', doc, re.S)
doc = doc.replace(heading.group(0), heading.group(1) + f'<w:bookmarkStart w:id="{bid}" w:name="_Appendix_A"/>'
                  + heading.group(2) + f'<w:bookmarkEnd w:id="{bid}"/>' + heading.group(3))
style = '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
links = (f'<w:hyperlink r:id="rIdLink1"><w:r>{style}<w:t>Project site</w:t></w:r></w:hyperlink>'
         '<w:r><w:t xml:space="preserve"> and </w:t></w:r>'
         f'<w:hyperlink w:anchor="_Appendix_A"><w:r>{style}<w:t>Appendix A</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r>')
start = doc.index("Further reading:")
end = doc.index("</w:p>", start)
doc = doc[:end] + links + doc[end:]
open(doc_path, "w", encoding="utf-8").write(doc)
open(rels_path, "w", encoding="utf-8").write(rels)
