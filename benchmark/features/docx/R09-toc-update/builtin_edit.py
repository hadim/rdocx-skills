"""Rewrite the TOC entries from the headings, with the page each heading lands on in LibreOffice's PDF."""
import re
import sys

from lxml import etree

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WNS
doc_path, pages_path = sys.argv[1], sys.argv[2]
pages = [{line.strip() for line in page.splitlines()} for page in open(pages_path).read().split("\f")]
tree = etree.parse(doc_path)
body = tree.getroot().find(W + "body")


def style(p):
    s = p.find(f"{W}pPr/{W}pStyle")
    return s.get(W + "val") if s is not None else ""


text = lambda p: "".join(t.text or "" for t in p.iter(W + "t"))  # noqa: E731
entries, page, next_id = [], 1, 9000
for p in body.iter(W + "p"):
    m = re.fullmatch(r"Heading([1-3])", style(p))
    if not m:
        continue
    page = next(i + 1 for i in range(page - 1, len(pages)) if text(p).strip() in pages[i])
    bm = p.find(W + "bookmarkStart")
    if bm is None or not bm.get(W + "name").startswith("_Toc"):  # a bookmark for the entry's link
        next_id += 1
        bm = etree.Element(W + "bookmarkStart", {W + "id": str(next_id), W + "name": f"_Toc{next_id}"})
        runs = p.findall(W + "r")
        runs[0].addprevious(bm)
        runs[-1].addnext(etree.Element(W + "bookmarkEnd", {W + "id": str(next_id)}))
    entries.append((int(m.group(1)), text(p), page, bm.get(W + "name")))
fld = lambda kind: f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'  # noqa: E731
paras = []
for k, (lvl, title, pg, bm) in enumerate(entries):
    start = (fld("begin") + '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
             + fld("separate")) if k == 0 else ""
    paras.append(
        f'<w:p xmlns:w="{WNS}"><w:pPr><w:pStyle w:val="TOC{lvl}"/><w:tabs><w:tab w:val="right" w:leader="dot" '
        f'w:pos="9350"/></w:tabs></w:pPr>{start}<w:hyperlink w:anchor="{bm}" w:history="1"><w:r><w:rPr><w:rStyle '
        f'w:val="Hyperlink"/></w:rPr><w:t>{title}</w:t></w:r><w:r><w:tab/></w:r>{fld("begin")}<w:r><w:instrText '
        f'xml:space="preserve"> PAGEREF {bm} \\h </w:instrText></w:r>{fld("separate")}<w:r><w:t>{pg}</w:t></w:r>'
        f'{fld("end")}</w:hyperlink></w:p>')
paras.append(f'<w:p xmlns:w="{WNS}">{fld("end")}</w:p>')
sdt = next(s for s in body.iter(W + "sdt") if s.find(f".//{W}docPartGallery[@{W}val='Table of Contents']") is not None)
content = sdt.find(W + "sdtContent")
for p in content.findall(W + "p")[1:]:  # keep the "Contents" heading, replace the entries
    content.remove(p)
for xml in paras:
    content.append(etree.fromstring(xml))
tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
