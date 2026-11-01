"""Deterministic, Word-like inputs for the authoring tasks D01-D14 (python-docx and lxml, run with common.PY).

Every make_input.py of a D task imports this module. The documents imitate what Word writes: paragraphs made
of several runs with different formatting, runs split in the middle of words with revision ids (w:rsidR),
proofing marks (w:proofErr) between them, a _GoBack bookmark. Text content is fixed, so that each task's
check.py can state what the output must hold.
"""
from __future__ import annotations

import datetime as dt
import struct
import zlib
from copy import deepcopy

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FIXED_TIME = dt.datetime(2026, 10, 1, 9, 0, 0)


def png_bytes(width: int, height: int, rgb: tuple[int, int, int], stripe=(255, 255, 255)) -> bytes:
    """A PNG with a diagonal stripe pattern (standard library only, deterministic)."""
    rows = []
    for y in range(height):
        row = bytearray(b"\x00")
        for x in range(width):
            row += bytes(stripe if (x + y) % 40 < 6 else rgb)
        rows.append(bytes(row))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b""))


_rsid_counter = [0x00A1B2C0]


def _rsid() -> str:
    _rsid_counter[0] += 0x111
    return f"{_rsid_counter[0]:08X}"


def new_document(title: str = "Benchmark input"):
    """A python-docx document set to A4 with 2.5 cm margins and fixed core properties."""
    _rsid_counter[0] = 0x00A1B2C0
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(s, side, Cm(2.5))
    cp = doc.core_properties
    cp.title, cp.author, cp.last_modified_by = title, "Benchmark", "Benchmark"
    cp.created = cp.modified = FIXED_TIME
    cp.revision = 3
    doc.settings.element.find(qn("w:zoom")).set(qn("w:percent"), "100")  # the template omits a required attribute
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    return doc


def add_runs(paragraph, parts):
    """Append runs to a paragraph. `parts` is a list of (text, flags) where flags is a string of
    b (bold), i (italic), u (underline), r (red colour), s (small, 9 pt); each run gets a revision id."""
    runs = []
    for text, flags in parts:
        run = paragraph.add_run(text)
        run.bold = True if "b" in flags else None
        run.italic = True if "i" in flags else None
        run.underline = True if "u" in flags else None
        if "r" in flags:
            run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        if "s" in flags:
            run.font.size = Pt(9)
        run._r.set(qn("w:rsidR"), _rsid())
        runs.append(run)
    paragraph._p.set(qn("w:rsidR"), _rsid())
    paragraph._p.set(qn("w:rsidRDefault"), _rsid())
    return runs


def para(doc_or_cell, parts, style=None):
    p = doc_or_cell.add_paragraph(style=style)
    add_runs(p, parts)
    return p


def proof_error_before(run, kind="spellStart"):
    """Insert a w:proofErr mark before a run, as Word's spell checker leaves them."""
    el = OxmlElement("w:proofErr")
    el.set(qn("w:type"), kind)
    run._r.addprevious(el)


def go_back_bookmark(paragraph):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), "0")
    start.set(qn("w:name"), "_GoBack")
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), "0")
    paragraph._p.append(start)
    paragraph._p.append(end)


def declare_rsids(doc):
    """List the revision ids used in the body in settings.xml (w:rsids), as Word does."""
    rsids = doc.settings.element.find(qn("w:rsids"))
    seen = {r.get(qn("w:val")) for r in rsids}
    for el in doc.element.body.iter():
        for attr in ("w:rsidR", "w:rsidRDefault", "w:rsidRPr"):
            v = el.get(qn(attr))
            if v and v not in seen:
                seen.add(v)
                r = OxmlElement("w:rsid")
                r.set(qn("w:val"), v)
                rsids.append(r)


COMMENT_STYLES = (
    f'<w:style {nsdecls("w")} w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>'
    '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
    '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>',
    f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/>'
    '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
    '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
)


def add_comment(doc, runs, text, author="Reviewer", initials="RV"):
    """python-docx's add_comment, plus the two styles it references and the template lacks."""
    if "CommentText" not in [s.style_id for s in doc.styles]:
        for xml in COMMENT_STYLES:
            doc.styles.element.append(parse_xml(xml))
    comment = doc.add_comment(runs, text=text, author=author, initials=initials)
    comment._comment_elm.set(qn("w:date"), "2026-10-01T09:00:00Z")  # python-docx stamps the current time
    return comment


FOOTNOTES_XML = (
    f'<w:footnotes {nsdecls("w", "r")}>'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '</w:footnotes>'
)

FOOTNOTE_STYLES = (
    f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>'
    '<w:basedOn w:val="Normal"/><w:link w:val="FootnoteTextChar"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
    f'<w:style {nsdecls("w")} w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar"><w:name w:val="Footnote Text Char"/>'
    '<w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="FootnoteText"/><w:uiPriority w:val="99"/>'
    '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
    f'<w:style {nsdecls("w")} w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
    '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
    '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
)


def add_footnote(doc, paragraph, text):
    """Add a footnote whose reference ends `paragraph` (python-docx has no API: package part + XML)."""
    main = doc.part
    part = next((r.target_part for r in main.rels.values() if r.reltype == RT.FOOTNOTES), None)
    if part is None:
        part = Part(PackURI("/word/footnotes.xml"),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                    FOOTNOTES_XML.encode(), main.package)
        main.relate_to(part, RT.FOOTNOTES)
        styles = doc.styles.element
        for xml in FOOTNOTE_STYLES:
            styles.append(parse_xml(xml))
    root = parse_xml(part.blob)
    fid = max(int(f.get(qn("w:id"))) for f in root.findall(qn("w:footnote"))) + 1
    root.append(parse_xml(
        f'<w:footnote {nsdecls("w")} w:id="{fid}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
        f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:footnote>'))
    from lxml import etree
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    paragraph._p.append(parse_xml(
        f'<w:r {nsdecls("w")}><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{fid}"/></w:r>'))
    return fid


EXTRA_NS = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
            'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:w10="urn:schemas-microsoft-com:office:word"')


def add_text_box(paragraph, text, box_id=1):
    """Anchor a Word text box (DrawingML shape with its VML fallback, as Word writes it) in `paragraph`."""
    content = (f'<w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent>')
    xml = (
        f'<w:r {nsdecls("w", "wp", "a")} {EXTRA_NS}>'
        '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3200400</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="2286000" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{100 + box_id}" name="Text Box {box_id}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        '<wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="457200"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln></wps:spPr>'
        f'<wps:txbx>{content}</wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">'
        '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
        '<mc:Fallback><w:pict>'
        f'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
        '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
        f'<v:shape id="Text Box {box_id}" o:spid="_x0000_s10{box_id}" type="#_x0000_t202" '
        'style="position:absolute;margin-left:252pt;margin-top:0;width:180pt;height:36pt;z-index:251659264">'
        f'<v:textbox>{content}</v:textbox><w10:wrap type="square"/></v:shape></w:pict></mc:Fallback>'
        '</mc:AlternateContent></w:r>'
    )
    paragraph._p.append(parse_xml(xml))


def clone_paragraph_after(paragraph):
    """A deep copy of a paragraph's XML inserted after it (returns the new lxml element)."""
    new = deepcopy(paragraph._p)
    paragraph._p.addnext(new)
    return new
