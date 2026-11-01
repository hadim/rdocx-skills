"""Deterministic, Word-like .docx inputs for the review tasks R01-R15 (input builders only).

The packages are written part by part, as Word 2016 writes them: a styles part with docDefaults and the
built-in styles the tasks use, settings with rsids, a theme, a font table, document and app properties,
w14:paraId / w14:textId on paragraphs, text fragmented into several runs (rsid splits, proofErr marks), and,
when a task needs them, comments (comments.xml, commentsExtended.xml, commentsIds.xml, people.xml),
tracked changes (w:ins / w:del with w:id, w:author, w:date), footnotes, headers and footers, numbering and
pictures. Every zip entry gets a fixed timestamp, so the same script writes the same bytes.
"""
from __future__ import annotations

import struct
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml"

NSDECL = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    f'xmlns:r="{R_NS}" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    f'xmlns:w="{W_NS}" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 w15 w16se w16cid wp14"'
)
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
FIXED_TIME = (2026, 10, 1, 9, 0, 0)

RSIDS = ["00A1B2C3", "00D4E5F6", "0017A9B0", "002C3D4E", "00E5F607", "005B6C7D", "00F1E2D3", "00C4B5A6"]


class Ids:
    """Deterministic w14:paraId values, rsids and revision ids."""

    def __init__(self, seed=0x1A2B0000):
        self.n = seed
        self.r = 0
        self.rev = 100

    def para(self):
        self.n += 0x111
        return f"{self.n & 0x7FFFFFFF:08X}"

    def rsid(self):
        self.r += 1
        return RSIDS[self.r % len(RSIDS)]

    def revision(self):
        self.rev += 1
        return self.rev


IDS = Ids()


def reset_ids():
    global IDS
    IDS = Ids()


# --------------------------------------------------------------------------------------------- runs
def rpr(bold=False, italic=False, color=None, size=None, font=None, lang=None, style=None, extra=""):
    parts = []
    if style:
        parts.append(f'<w:rStyle w:val="{style}"/>')
    if font:
        parts.append(f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>')
    if bold:
        parts.append("<w:b/><w:bCs/>")
    if italic:
        parts.append("<w:i/><w:iCs/>")
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    if size:
        parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    if lang:
        parts.append(f'<w:lang w:val="{lang}"/>')
    parts.append(extra)
    body = "".join(parts)
    return f"<w:rPr>{body}</w:rPr>" if body else ""


def t(text, tag="w:t"):
    sp = ' xml:space="preserve"' if text != text.strip() or "  " in text else ""
    return f"<{tag}{sp}>{escape(text)}</{tag}>"


def run(text, props="", deleted=False):
    """One run. `props` is an rPr string (rpr(...))."""
    return f'<w:r w:rsidR="{IDS.rsid()}">{props}{t(text, "w:delText" if deleted else "w:t")}</w:r>'


def frag(text, cuts=(), props="", spell=()):
    """Runs of `text` split at the character offsets `cuts`, as Word leaves text after edits on different
    days (one rsid per run); `spell` lists (start, end) offsets wrapped in proofErr spellStart / spellEnd."""
    bounds = sorted(set([0, len(text), *cuts, *[o for s in spell for o in s]]))
    out = []
    for a, b in zip(bounds, bounds[1:]):
        if any(a == s for s, _ in spell):
            out.append('<w:proofErr w:type="spellStart"/>')
        out.append(run(text[a:b], props))
        if any(b == e for _, e in spell):
            out.append('<w:proofErr w:type="spellEnd"/>')
    return "".join(out)


def ins(content, author, date, rid=None):
    return f'<w:ins w:id="{rid or IDS.revision()}" w:author="{escape(author)}" w:date="{date}">{content}</w:ins>'


def dele(text, author, date, props="", rid=None):
    return (f'<w:del w:id="{rid or IDS.revision()}" w:author="{escape(author)}" w:date="{date}">'
            f"{run(text, props, deleted=True)}</w:del>")


def rpr_change(new_props_inner, old_props_inner, author, date):
    """rPr of a run whose formatting was changed with tracking on: current props, then the old ones."""
    return (f"<w:rPr>{new_props_inner}<w:rPrChange w:id=\"{IDS.revision()}\" w:author=\"{escape(author)}\" "
            f"w:date=\"{date}\"><w:rPr>{old_props_inner}</w:rPr></w:rPrChange></w:rPr>")


# ---------------------------------------------------------------------------------------- paragraphs
def para(content="", style=None, ppr="", mark_rpr="", page_break_before=False):
    """A paragraph with Word's paraId / textId / rsid attributes. `mark_rpr` goes in pPr/rPr (the
    paragraph mark's run properties: <w:ins .../> or <w:del .../> for a tracked paragraph mark)."""
    pid = IDS.para()
    inner = ""
    if style:
        inner += f'<w:pStyle w:val="{style}"/>'
    if page_break_before:
        inner += "<w:pageBreakBefore/>"
    inner += ppr
    if mark_rpr:
        inner += f"<w:rPr>{mark_rpr}</w:rPr>"
    ppr_xml = f"<w:pPr>{inner}</w:pPr>" if inner else ""
    rs = IDS.rsid()
    return (f'<w:p w14:paraId="{pid}" w14:textId="77777777" w:rsidR="{rs}" w:rsidRDefault="{rs}" '
            f'w:rsidP="{rs}">{ppr_xml}{content}</w:p>')


def mark(kind, author, date):
    """A tracked paragraph mark (kind 'ins' or 'del') for para(mark_rpr=...)."""
    return f'<w:{kind} w:id="{IDS.revision()}" w:author="{escape(author)}" w:date="{date}"/>'


def page_break():
    return f'<w:r w:rsidR="{IDS.rsid()}"><w:br w:type="page"/></w:r>'


def table(rows, widths=None, style="TableGrid", header=True):
    """rows: list of lists of cell contents (each a paragraph XML string or plain text)."""
    ncol = len(rows[0])
    widths = widths or [9360 // ncol] * ncol
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out = [f'<w:tbl><w:tblPr><w:tblStyle w:val="{style}"/><w:tblW w:w="0" w:type="auto"/>'
           f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
           f'w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>']
    for ri, row in enumerate(rows):
        out.append(f'<w:tr w:rsidR="{IDS.rsid()}" w14:paraId="{IDS.para()}" w14:textId="77777777">')
        for ci, cell in enumerate(row):
            content = cell if cell.startswith("<w:p ") else para(frag(cell, props=rpr(bold=header and ri == 0)))
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/></w:tcPr>{content}</w:tc>')
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def field(instr, result, props=""):
    """A complex field as Word writes it: begin, instrText, separate, result run, end."""
    r = IDS.rsid
    return (f'<w:r w:rsidR="{r()}">{props}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r w:rsidR="{r()}">{props}<w:instrText xml:space="preserve"> {escape(instr)} </w:instrText></w:r>'
            f'<w:r w:rsidR="{r()}">{props}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r w:rsidR="{r()}">{props}{t(result)}</w:r>'
            f'<w:r w:rsidR="{r()}">{props}<w:fldChar w:fldCharType="end"/></w:r>')


def bookmark(bid, name, content):
    return f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>{content}<w:bookmarkEnd w:id="{bid}"/>'


def comment_start(cid):
    return f'<w:commentRangeStart w:id="{cid}"/>'


def comment_end(cid):
    return (f'<w:commentRangeEnd w:id="{cid}"/><w:r w:rsidR="{IDS.rsid()}"><w:rPr><w:rStyle '
            f'w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>')


def footnote_ref(fid):
    return (f'<w:r w:rsidR="{IDS.rsid()}"><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
            f'<w:footnoteReference w:id="{fid}"/></w:r>')


def sect_pr(header_rid=None, footer_rid=None):
    refs = ""
    if header_rid:
        refs += f'<w:headerReference w:type="default" r:id="{header_rid}"/>'
    if footer_rid:
        refs += f'<w:footerReference w:type="default" r:id="{footer_rid}"/>'
    return (f'<w:sectPr w:rsidR="{RSIDS[0]}" w:rsidSect="{RSIDS[0]}">{refs}<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" '
            'w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')


# ------------------------------------------------------------------------------------------ pictures
def png(width, height, rgb):
    """A small PNG of one colour with a darker diagonal band (standard library only)."""
    rows = []
    for y in range(height):
        row = bytearray(b"\x00")
        for x in range(width):
            band = abs(x * height - y * width) < width * 3
            row += bytes(c // 2 for c in rgb) if band else bytes(rgb)
        rows.append(bytes(row))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b""))


def inline_picture(rid, cx, cy, pic_id, name):
    return (
        f'<w:r w:rsidR="{IDS.rsid()}"><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="{IDS.para()}" wp14:editId="{IDS.para()}">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{pic_id}" name="{name}"/><wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        '</wp:inline></w:drawing></w:r>')


# ------------------------------------------------------------------------------------------- styles
def _pstyle(sid, name, based="Normal", nxt=None, ppr="", rpr_="", extra="", link=None):
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>'
            + (f'<w:basedOn w:val="{based}"/>' if based else "")
            + (f'<w:next w:val="{nxt}"/>' if nxt else "")
            + (f'<w:link w:val="{link}"/>' if link else "")
            + extra
            + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + (f"<w:rPr>{rpr_}</w:rPr>" if rpr_ else "") + "</w:style>")


def _cstyle(sid, name, rpr_="", link=None, extra=""):
    return (f'<w:style w:type="character" w:styleId="{sid}"><w:name w:val="{name}"/>'
            '<w:basedOn w:val="DefaultParagraphFont"/>' + (f'<w:link w:val="{link}"/>' if link else "")
            + extra + (f"<w:rPr>{rpr_}</w:rPr>" if rpr_ else "") + "</w:style>")


def styles_xml(heading_font="Calibri Light", heading_color="2F5496", body_font="Calibri", extra_styles=""):
    hf = f'<w:rFonts w:ascii="{heading_font}" w:hAnsi="{heading_font}" w:eastAsia="{heading_font}" w:cs="Times New Roman"/>'
    heading = lambda lvl, size, before, color: _pstyle(  # noqa: E731
        f"Heading{lvl}", f"heading {lvl}", nxt="Normal", link=f"Heading{lvl}Char",
        extra='<w:uiPriority w:val="9"/><w:qFormat/>',
        ppr=f'<w:keepNext/><w:keepLines/><w:spacing w:before="{before}" w:after="0"/><w:outlineLvl w:val="{lvl - 1}"/>',
        rpr_=f'{hf}<w:color w:val="{color}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    s = [
        f'<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="{R_NS}" '
        f'xmlns:w="{W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">',
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{body_font}" w:eastAsia="{body_font}" w:hAnsi="{body_font}" w:cs="Times New Roman"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
        '</w:pPr></w:pPrDefault></w:docDefaults>',
        '<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" '
        'w:defQFormat="0" w:count="376"><w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>'
        '<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/></w:latentStyles>',
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>',
        heading(1, 32, 240, heading_color), heading(2, 26, 40, heading_color), heading(3, 24, 40, "1F3763"),
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default '
        'Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" '
        'w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
        *[_cstyle(f"Heading{i}Char", f"Heading {i} Char", link=f"Heading{i}",
                  rpr_=f'{hf}<w:color w:val="{heading_color}"/>') for i in (1, 2, 3)],
        _pstyle("Title", "Title", nxt="Normal", link="TitleChar", extra='<w:uiPriority w:val="10"/><w:qFormat/>',
                ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
                rpr_=f'{hf}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>'),
        _cstyle("TitleChar", "Title Char", link="Title", rpr_=f'{hf}<w:sz w:val="56"/>'),
        _pstyle("TOCHeading", "TOC Heading", based="Heading1", nxt="Normal",
                extra='<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>', ppr='<w:outlineLvl w:val="9"/>'),
        _pstyle("TOC1", "toc 1", nxt="Normal", extra='<w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>',
                ppr='<w:spacing w:after="100"/>'),
        _pstyle("TOC2", "toc 2", nxt="Normal", extra='<w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>',
                ppr='<w:spacing w:after="100"/><w:ind w:left="220"/>'),
        _pstyle("TOC3", "toc 3", nxt="Normal", extra='<w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>',
                ppr='<w:spacing w:after="100"/><w:ind w:left="440"/>'),
        _cstyle("Hyperlink", "Hyperlink", rpr_='<w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/>',
                extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
        _cstyle("CommentReference", "annotation reference", rpr_='<w:sz w:val="16"/><w:szCs w:val="16"/>',
                extra='<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'),
        _pstyle("CommentText", "annotation text", link="CommentTextChar",
                extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>', ppr='<w:spacing w:line="240" w:lineRule="auto"/>',
                rpr_='<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        _cstyle("CommentTextChar", "Comment Text Char", link="CommentText", rpr_='<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        _pstyle("CommentSubject", "annotation subject", based="CommentText", nxt="CommentText",
                extra='<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>', rpr_="<w:b/><w:bCs/>"),
        _pstyle("FootnoteText", "footnote text", link="FootnoteTextChar",
                extra='<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>',
                ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rpr_='<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        _cstyle("FootnoteTextChar", "Footnote Text Char", link="FootnoteText", rpr_='<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        _cstyle("FootnoteReference", "footnote reference", rpr_='<w:vertAlign w:val="superscript"/>',
                extra='<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'),
        _pstyle("Header", "header", link="HeaderChar", extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>',
                ppr='<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>'
                    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'),
        _cstyle("HeaderChar", "Header Char", link="Header"),
        _pstyle("Footer", "footer", link="FooterChar", extra='<w:uiPriority w:val="99"/><w:unhideWhenUsed/>',
                ppr='<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>'
                    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'),
        _cstyle("FooterChar", "Footer Char", link="Footer"),
        _pstyle("ListParagraph", "List Paragraph", extra='<w:uiPriority w:val="34"/><w:qFormat/>',
                ppr='<w:ind w:left="720"/><w:contextualSpacing/>'),
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>'
        '<w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr>'
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" '
        'w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right '
        'w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" '
        'w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
        '</w:style>',
        extra_styles,
        "</w:styles>",
    ]
    return XML_HEAD + "".join(s)


def numbering_xml(abstracts):
    """abstracts: list of (abstract_id, num_id, fmt, lvl_text, left, hanging)."""
    body = [f'<w:numbering {NSDECL}>']
    for aid, _num, fmt, text, left, hang in abstracts:
        body.append(f'<w:abstractNum w:abstractNumId="{aid}" w15:restartNumberingAfterBreak="0">'
                    f'<w:nsid w:val="{IDS.para()}"/><w:multiLevelType w:val="hybridMultilevel"/>'
                    f'<w:tmpl w:val="{0x0409000F + aid:08X}"/>'
                    f'<w:lvl w:ilvl="0" w:tplc="0409000F"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                    f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" '
                    f'w:hanging="{hang}"/></w:pPr></w:lvl></w:abstractNum>')
    for aid, num, *_ in abstracts:
        body.append(f'<w:num w:numId="{num}"><w:abstractNumId w:val="{aid}"/></w:num>')
    body.append("</w:numbering>")
    return XML_HEAD + "".join(body)


# ------------------------------------------------------------------------------------------ package
THEME = None


def _theme():
    global THEME
    if THEME is None:
        import docx  # the theme part of python-docx's default template (Office theme), read at build time
        tpl = Path(docx.__file__).parent / "templates" / "default.docx"
        with zipfile.ZipFile(tpl) as z:
            THEME = z.read("word/theme/theme1.xml")
    return THEME


class Package:
    """An OPC package built part by part, saved with fixed zip timestamps."""

    def __init__(self, title="Service agreement", author="Alex Author"):
        self.parts: dict[str, bytes] = {}
        self.ctypes: dict[str, str] = {}
        self.rels: list[tuple[str, str, str]] = []
        self.styles = styles_xml()
        self.settings_extra = ""
        self.title, self.author = title, author

    def add(self, name, data, ctype=None, rel_type=None):
        """Add a part under word/ (name relative to word/); returns the relationship id if related."""
        self.parts["word/" + name] = data.encode() if isinstance(data, str) else data
        if ctype:
            self.ctypes["/word/" + name] = ctype
        if rel_type:
            rid = f"rId{len(self.rels) + 1}"
            self.rels.append((rid, rel_type, name))
            return rid
        return None

    def reserve(self, name, ctype, rel_type):
        """Relate a part whose XML is set later (headers, footers)."""
        return self.add(name, b"", ctype, rel_type)

    def save(self, path, body):
        p = self.parts
        self.add("styles.xml", self.styles, f"{CT_W}.styles+xml", f"{REL}/styles")
        settings = (
            f'<w:settings {NSDECL}><w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/>'
            '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>'
            + self.settings_extra +
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
            'w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting '
            'w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" '
            'w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:rsids><w:rsidRoot '
            f'w:val="{RSIDS[0]}"/>' + "".join(f'<w:rsid w:val="{r}"/>' for r in sorted(RSIDS)) + '</w:rsids>'
            '<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/>'
            '<m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc '
            'm:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim '
            'm:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1" '
            'w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" '
            'w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" '
            'w:followedHyperlink="followedHyperlink"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>'
            '<w14:docId w14:val="5C2B1A9E"/><w15:chartTrackingRefBased/>'
            '<w15:docId w15:val="{7B6F3C2A-1D4E-4B8F-9A0C-2E5D6F7A8B9C}"/></w:settings>')
        self.add("settings.xml", XML_HEAD + settings, f"{CT_W}.settings+xml", f"{REL}/settings")
        self.add("webSettings.xml", XML_HEAD + f'<w:webSettings {NSDECL}><w:optimizeForBrowser/>'
                 '<w:allowPNG/></w:webSettings>', f"{CT_W}.webSettings+xml", f"{REL}/webSettings")
        fonts = "".join(
            f'<w:font w:name="{n}"><w:panose1 w:val="{pan}"/><w:charset w:val="00"/><w:family w:val="{fam}"/>'
            '<w:pitch w:val="variable"/></w:font>'
            for n, pan, fam in [("Calibri", "020F0502020204030204", "swiss"),
                                ("Times New Roman", "02020603050405020304", "roman"),
                                ("Calibri Light", "020F0302020204030204", "swiss"),
                                ("Georgia", "02040502050405020303", "roman")])
        self.add("fontTable.xml", XML_HEAD + f'<w:fonts {NSDECL}>{fonts}</w:fonts>',
                 f"{CT_W}.fontTable+xml", f"{REL}/fontTable")
        self.add("theme/theme1.xml", _theme(), "application/vnd.openxmlformats-officedocument.theme+xml",
                 f"{REL}/theme")
        doc = XML_HEAD + f"<w:document {NSDECL}><w:body>{body}</w:body></w:document>"
        p["word/document.xml"] = doc.encode()
        ct = ['<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
              '<Default Extension="png" ContentType="image/png"/>',
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
              '<Default Extension="xml" ContentType="application/xml"/>',
              f'<Override PartName="/word/document.xml" ContentType="{CT_W}.document.main+xml"/>']
        ct += [f'<Override PartName="{k}" ContentType="{v}"/>' for k, v in self.ctypes.items()]
        ct += ['<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.'
               'core-properties+xml"/>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.'
               'openxmlformats-officedocument.extended-properties+xml"/></Types>']
        rels = "".join(f'<Relationship Id="{i}" Type="{ty}" Target="{tg}"/>' for i, ty, tg in self.rels)
        core = (f'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:title>{escape(self.title)}</dc:title><dc:creator>{escape(self.author)}</dc:creator>'
                f'<cp:lastModifiedBy>{escape(self.author)}</cp:lastModifiedBy><cp:revision>3</cp:revision>'
                '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-01T08:00:00Z</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-01T09:00:00Z</dcterms:modified>'
                '</cp:coreProperties>')
        app = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
               'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm'
               '</Template><TotalTime>42</TotalTime><Pages>1</Pages><Application>Microsoft Office Word</Application>'
               '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false'
               '</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>'
               '<AppVersion>16.0000</AppVersion></Properties>')
        entries = [
            ("[Content_Types].xml", XML_HEAD + "".join(ct)),
            ("_rels/.rels", XML_HEAD + f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId3" Type="'
             'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
             'Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/'
             '2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" '
             f'Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'),
            ("word/document.xml", p["word/document.xml"]),
            ("word/_rels/document.xml.rels", XML_HEAD + f'<Relationships xmlns="{PKG_REL}">{rels}</Relationships>'),
        ]
        entries += [(k, v) for k, v in sorted(p.items()) if k != "word/document.xml"]
        entries += [("docProps/core.xml", XML_HEAD + core), ("docProps/app.xml", XML_HEAD + app)]
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in entries:
                info = zipfile.ZipInfo(name, FIXED_TIME)
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, data.encode() if isinstance(data, str) else data)
        return path

    # ---- optional parts ------------------------------------------------------------------------
    def comments(self, comments):
        """comments: list of dicts id, author, initials, date, paragraphs (list of str), parent (id or None),
        done (bool). Writes comments.xml, commentsExtended.xml, commentsIds.xml and people.xml as Word does;
        the paraId of a comment is the one of its last paragraph."""
        xml, ext, ids, last = [], [], [], {}
        for c in comments:
            ps = []
            for k, text in enumerate(c["paragraphs"]):
                pid = IDS.para()
                ref = ('<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>'
                       if k == 0 else "")
                ps.append(f'<w:p w14:paraId="{pid}" w14:textId="77777777" w:rsidR="{RSIDS[3]}" '
                          f'w:rsidRDefault="{RSIDS[3]}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>{ref}'
                          f'<w:r>{t(text)}</w:r></w:p>')
            last[c["id"]] = pid
            xml.append(f'<w:comment w:id="{c["id"]}" w:author="{escape(c["author"])}" w:date="{c["date"]}" '
                       f'w:initials="{c["initials"]}">{"".join(ps)}</w:comment>')
            parent = f' w15:paraIdParent="{last[c["parent"]]}"' if c.get("parent") is not None else ""
            ext.append(f'<w15:commentEx w15:paraId="{pid}"{parent} w15:done="{1 if c.get("done") else 0}"/>')
            ids.append(f'<w16cid:commentId w16cid:paraId="{pid}" w16cid:durableId="{0x3C4D5E00 + c["id"]:08X}"/>')
        self.add("comments.xml", XML_HEAD + f'<w:comments {NSDECL}>{"".join(xml)}</w:comments>',
                 f"{CT_W}.comments+xml", f"{REL}/comments")
        self.add("commentsExtended.xml", XML_HEAD + f'<w15:commentsEx {NSDECL}>{"".join(ext)}</w15:commentsEx>',
                 f"{CT_W}.commentsExtended+xml", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
        self.add("commentsIds.xml", XML_HEAD + f'<w16cid:commentsIds {NSDECL}>{"".join(ids)}</w16cid:commentsIds>',
                 f"{CT_W}.commentsIds+xml", "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds")
        people = "".join(
            f'<w15:person w15:author="{escape(a)}"><w15:presenceInfo w15:providerId="None" w15:userId="{escape(a)}"/>'
            "</w15:person>" for a in dict.fromkeys(c["author"] for c in comments))
        self.add("people.xml", XML_HEAD + f'<w15:people {NSDECL}>{people}</w15:people>',
                 f"{CT_W}.people+xml", "http://schemas.microsoft.com/office/2011/relationships/people")

    def notes(self, footnotes):
        """footnotes: {id: paragraph content XML}. Writes footnotes.xml and endnotes.xml with the separators."""
        sep = lambda kind, tag: (  # noqa: E731
            f'<w:{kind} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
            f'w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:{kind}>'
            f'<w:{kind} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" '
            f'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>')
        fn = "".join(
            f'<w:footnote w:id="{i}"><w:p w14:paraId="{IDS.para()}" w14:textId="77777777"><w:pPr><w:pStyle '
            'w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/>'
            f'</w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>{content}</w:p></w:footnote>'
            for i, content in footnotes.items())
        self.add("footnotes.xml", XML_HEAD + f'<w:footnotes {NSDECL}>{sep("footnote", "f")}{fn}</w:footnotes>',
                 f"{CT_W}.footnotes+xml", f"{REL}/footnotes")
        self.add("endnotes.xml", XML_HEAD + f'<w:endnotes {NSDECL}>{sep("endnote", "e")}</w:endnotes>',
                 f"{CT_W}.endnotes+xml", f"{REL}/endnotes")
        self.settings_extra += ('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
                                '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>')

    def header_footer(self, header_content=None, footer_content=None):
        """Default header and footer parts; returns (header rid, footer rid) for sect_pr()."""
        hr = fr = None
        if header_content is not None:
            hr = self.add("header1.xml", XML_HEAD + f'<w:hdr {NSDECL}>{header_content}</w:hdr>',
                          f"{CT_W}.header+xml", f"{REL}/header")
        if footer_content is not None:
            fr = self.add("footer1.xml", XML_HEAD + f'<w:ftr {NSDECL}>{footer_content}</w:ftr>',
                          f"{CT_W}.footer+xml", f"{REL}/footer")
        return hr, fr

    def numbering(self, abstracts):
        self.add("numbering.xml", numbering_xml(abstracts), f"{CT_W}.numbering+xml", f"{REL}/numbering")

    def image(self, name, data):
        return self.add(f"media/{name}", data, None, f"{REL}/image")


def numbered(content, num_id, style="ListParagraph"):
    return para(content, style=style, ppr=f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')


def lorem(seed, words=60):
    """Deterministic filler prose for page-filling paragraphs."""
    vocab = ("the supplier shall keep records of every delivery and make them available to the customer on "
             "request within a reasonable time after each inspection of the premises and equipment used for "
             "the services described in this agreement including any changes agreed in writing by both parties "
             "under the procedure set out below").split()
    out, x = [], seed * 7919 + 13
    for _ in range(words):
        x = (x * 1103515245 + 12345) & 0x7FFFFFFF
        out.append(vocab[x % len(vocab)])
    s = " ".join(out)
    return s[0].upper() + s[1:] + "."
