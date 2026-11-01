# bench: tools=python-pptx (OPC parts), lxml; processes=1; raw_xml=yes
import sys

from pptx import Presentation
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml

P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
DANA, SAM = "{6F1B3A52-0000-4000-8000-000000000001}", "{6F1B3A52-0000-4000-8000-000000000002}"
prs = Presentation(sys.argv[1])
slide = prs.slides[2]
key = next(sh for sh in slide.shapes if sh.name == "Key figure")
authors = (f'<p188:authorLst xmlns:p188="{P188}">'
           f'<p188:author id="{DANA}" name="Dana Reviewer" initials="DR" userId="Dana Reviewer" providerId="None"/>'
           f'<p188:author id="{SAM}" name="Sam Author" initials="SA" userId="Sam Author" providerId="None"/>'
           f'</p188:authorLst>')
comments = (
    f'<p188:cmLst xmlns:p188="{P188}" xmlns:a="{A}">'
    f'<p188:cm id="{{6F1B3A52-0000-4000-8000-0000000000A1}}" authorId="{DANA}" created="2026-10-01T12:00:00Z">'
    f'<ac:deMkLst xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command" '
    f'xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command">'
    f'<pc:docMk/><pc:sldMk cId="0" sldId="{slide.slide_id}"/><ac:spMk id="{key.shape_id}"/></ac:deMkLst>'
    f'<p188:replyLst><p188:reply id="{{6F1B3A52-0000-4000-8000-0000000000A2}}" authorId="{SAM}" '
    f'created="2026-10-01T12:05:00Z"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
    f'<a:t>Updated with Q3 numbers.</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>'
    f'<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Please check this figure.</a:t>'
    f'</a:r></a:p></p188:txBody></p188:cm></p188:cmLst>')
pkg = prs.part.package
authors_part = Part(PackURI("/ppt/authors.xml"), "application/vnd.ms-powerpoint.authors+xml", pkg, authors.encode())
prs.part.relate_to(authors_part, "http://schemas.microsoft.com/office/2018/10/relationships/authors")
cm_part = Part(PackURI(f"/ppt/comments/modernComment_{slide.slide_id:X}.xml"),
               "application/vnd.ms-powerpoint.comments+xml", pkg, comments.encode())
rid = slide.part.relate_to(cm_part, "http://schemas.microsoft.com/office/2018/10/relationships/comments")
slide._element.append(parse_xml(  # PowerPoint reads a slide's modern comments through this extension
    f'<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ext '
    f'uri="{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}"><p188:commentRel xmlns:p188="{P188}" xmlns:r='
    f'"http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/></p:ext></p:extLst>'))
prs.save(sys.argv[2])
