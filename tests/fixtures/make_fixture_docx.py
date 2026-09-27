"""Generate fixture-report.docx: a synthetic survey report of 13 pages in LibreOffice (12 in rdocx on 9a7ed714),
written from scratch (invented text, drawn pictures), with the serialisation traits of a file exported by
Google Docs and then edited in Word.

Usage: make_fixture_docx.py [OUT] [TOC_PAGES]
  OUT        output path (default fixture-report.docx); the output is deterministic
  TOC_PAGES  comma-separated page numbers cached in the table of contents entries (default: the pages
             LibreOffice gives, so that the default output is the attached fixture)

Every trait below is injected on purpose, and each one is a row or a column of the acceptance matrices:
- identity attributes: w:rsidR, w:rsidRDefault and w:rsidP on paragraphs; w:rsidR and w:rsidRPr on runs,
  field runs included; w14:paraId and w14:textId on paragraphs and table rows; w:rsidR and w:rsidTr on
  table rows; the matching w:rsids list in settings.xml
- the table of contents in a block content control (docPartObj, w:id), its TOC field packed in one run
  (begin, instruction, separate) and its cached entries written as plain text
- PAGE and NUMPAGES in the footer, each packed in one run with no cached result
- two Google Docs content controls, each with a w:id: one inline around a run (tag goog_rdk_0), one at
  body level around a paragraph (tag goog_rdk_1)
- xml:space="preserve" on every w:t; explicit w:val="0" toggles (w:b, w:i, w:rtl, w:keepNext,
  w:pageBreakBefore); w:orient="portrait"; a default namespace on every part root
- document defaults at line 276 auto with Calibri body text and Cambria headings
- a 56-row table crossing pages with a repeated header row and identity attributes on every row; a table
  with a merged total row
- a tall inline picture followed by its caption, with no keep-with-next; a wide picture; a cover picture
- external hyperlinks to example.org; an empty comments part; a Google Docs custom XML part
"""
import io, os, random, re, sys, zipfile
from xml.sax.saxutils import escape
from PIL import Image, ImageDraw

OUT = sys.argv[1] if len(sys.argv) > 1 else "fixture-report.docx"
TOC_PAGES = [int(x) for x in (sys.argv[2] if len(sys.argv) > 2 else "3,3,3,4,5,6,9,9,9,9,10,10,10,10,11,11,11,12,12,12,13").split(",")]
R = random.Random(1907)
EMU_PT = 12700

# ---------------------------------------------------------------- namespaces and identity attributes
NS = ('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word" '
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
      'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
      'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
      'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" '
      'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
      'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
      'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" '
      'mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"')
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
RSIDS = ["00A31F2C", "00B7440E", "003C19D5", "00E6522B", "0061A7F0", "00D90C13", "004F2E81"]
_para_ids = set()


def rsid():
    return R.choice(RSIDS)


def para_id():
    while True:
        v = f"{R.randrange(0x10000000, 0x7FFFFFFF):08X}"
        if v not in _para_ids:
            _para_ids.add(v)
            return v


def p_attrs():
    text_id = "77777777" if R.random() < 0.6 else para_id()
    return (f'w:rsidR="{rsid()}" w:rsidRDefault="{rsid()}" w:rsidP="{rsid()}" '
            f'w14:paraId="{para_id()}" w14:textId="{text_id}"')


def t(s):
    return f'<w:t xml:space="preserve">{escape(s)}</w:t>'


# ---------------------------------------------------------------- runs and paragraphs
RPR_ORDER = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike vanish color spacing w kern position sz szCs "
             "highlight u effect shd vertAlign rtl cs lang").split()


def rpr_sorted(fragment):
    """Run properties in schema order (every child written here is an empty element)."""
    items = re.findall(r"<w:(\w+)[^>]*/>", fragment)
    parts = re.findall(r"<w:\w+[^>]*/>", fragment)
    return "".join(p for _, p in sorted(zip(items, parts), key=lambda x: RPR_ORDER.index(x[0])))


def run(text="", rpr="", toggles=True, inner=None):
    """A run with identity attributes; Google Docs style w:val="0" toggles on most runs."""
    props = rpr
    if toggles:
        props += '<w:b w:val="0"/><w:i w:val="0"/><w:rtl w:val="0"/>' if not ("<w:b/>" in rpr or "<w:i/>" in rpr) else '<w:rtl w:val="0"/>'
    props = rpr_sorted(props)
    body = inner if inner is not None else t(text)
    rp = f"<w:rPr>{props}</w:rPr>" if props else ""
    attrs = f'w:rsidR="{rsid()}"' + (f' w:rsidRPr="{rsid()}"' if R.random() < 0.5 else "")
    return f"<w:r {attrs}>{rp}{body}</w:r>"


def para(content, style=None, ppr=""):
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    pp = f"<w:pPr>{st}{ppr}</w:pPr>" if (st or ppr) else ""
    return f"<w:p {p_attrs()}>{pp}{content}</w:p>"


def text_para(text, style=None, ppr="", google=True):
    extra = '<w:keepNext w:val="0"/><w:pageBreakBefore w:val="0"/>' if google and style is None else ""
    return para(run(text, toggles=google), style, extra + ppr)


def heading(level, text, ppr=""):
    return para(run(text, toggles=False), f"Heading{level}", ppr)


def bullet(text, num_id=1):
    return para(run(text), "ListParagraph", f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')


def page_break():
    return para(run(inner='<w:br w:type="page"/>', toggles=False))


def hyperlink(rid, text):
    return (f'<w:hyperlink r:id="{rid}" w:history="1">'
            + run(text, '<w:rStyle w:val="Hyperlink"/>', toggles=False) + "</w:hyperlink>")


# ---------------------------------------------------------------- pictures
RELS = []          # (rid, type, target, external)
MEDIA = {}         # name -> bytes
_doc_pr = [0]


def add_rel(kind, target, external=False):
    rid = f"rId{len(RELS) + 1}"
    RELS.append((rid, kind, target, external))
    return rid


def picture(img, width_pt, height_pt, descr):
    name = f"image{len(MEDIA) + 1}.png"
    buf = io.BytesIO(); img.save(buf, "PNG"); MEDIA[name] = buf.getvalue()
    rid = add_rel("image", f"media/{name}")
    _doc_pr[0] += 1
    n = _doc_pr[0]
    cx, cy = int(width_pt * EMU_PT), int(height_pt * EMU_PT)
    return (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{n}" name="Picture {n}" descr="{escape(descr)}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="{n}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing>')


def img_cover():
    im = Image.new("RGB", (600, 360), (244, 239, 230)); d = ImageDraw.Draw(im)
    d.rectangle([0, 260, 600, 360], fill=(96, 128, 150))
    d.arc([60, 80, 540, 440], 180, 360, fill=(123, 30, 58), width=14)
    d.line([30, 262, 570, 262], fill=(62, 74, 89), width=10)
    for x in range(132, 470, 42):
        top = 260 - int((240 ** 2 - (x - 300) ** 2) ** 0.5 * 0.75)
        d.line([x, 262, x, top], fill=(168, 107, 18), width=3)
    return im


def img_section():
    im = Image.new("RGB", (700, 1000), (255, 255, 255)); d = ImageDraw.Draw(im)
    d.rectangle([40, 120, 660, 180], fill=(190, 190, 195), outline=(62, 74, 89), width=4)      # deck
    d.rectangle([60, 90, 80, 120], fill=(123, 30, 58)); d.rectangle([620, 90, 640, 120], fill=(123, 30, 58))
    d.polygon([(250, 180), (450, 180), (420, 760), (280, 760)], fill=(214, 206, 190), outline=(62, 74, 89))
    d.rectangle([180, 760, 520, 860], fill=(168, 107, 18), outline=(62, 74, 89), width=4)     # footing
    d.line([0, 560, 700, 560], fill=(96, 128, 150), width=6)                                    # water line
    for y in range(880, 1000, 24):
        d.line([100, y, 600, y], fill=(200, 200, 200), width=2)
    for x in (300, 350, 400):
        d.line([x, 860, x, 990], fill=(62, 74, 89), width=5)                                     # piles
    return im


def img_plan():
    im = Image.new("RGB", (1200, 480), (250, 250, 247)); d = ImageDraw.Draw(im)
    d.polygon([(0, 180), (1200, 90), (1200, 300), (0, 390)], fill=(170, 196, 214))
    d.line([420, 60, 560, 430], fill=(123, 30, 58), width=12)                                  # footbridge
    d.line([0, 460, 1200, 420], fill=(150, 150, 150), width=18)                                 # road
    d.line([0, 40, 1200, 20], fill=(150, 150, 150), width=18)
    d.polygon([(1120, 140), (1100, 200), (1140, 200)], fill=(62, 74, 89))                       # north arrow
    for x in range(40, 1200, 110):
        d.ellipse([x, 400 - (x % 3) * 8, x + 26, 426 - (x % 3) * 8], fill=(92, 127, 103))       # trees
    return im


# ---------------------------------------------------------------- invented text
ELEMENTS = ["deck soffit", "parapet", "handrail", "bearing", "abutment", "wing wall", "drainage outlet",
            "expansion joint", "lighting column", "pier", "surfacing", "balustrade"]
LOCATIONS = ["span 1, north edge", "span 1, south edge", "span 2, midspan", "span 2, quarter point",
             "pier 1, west face", "pier 2, east face", "west abutment", "east abutment", "span 3, south edge"]
OBS = ["hairline cracking along the soffit, about {n} m long", "surface corrosion on {n} per cent of the exposed steel",
       "loose coping stones over {n} m", "a blocked outlet with standing water after rain",
       "missing sealant at the joint over {n} mm", "minor spalling exposing the reinforcement over {n} cm",
       "paint breakdown at the base plates", "vegetation growing in the joints",
       "worn anti-slip surfacing over {n} square metres", "graffiti on {n} parapet panels",
       "a bearing plate displaced by about {n} mm", "efflorescence below the deck joints",
       "a damaged lamp cover", "split decking boards over {n} m"]
ACTIONS = ["monitor", "clean and reseal", "repaint", "replace", "repair within 12 months", "urgent repair", "no action"]
RATING_WORDS = {1: "very good", 2: "good", 3: "fair", 4: "poor", 5: "very poor"}
FILLER = [
    "Access to the soffit was gained from a small boat on the second day of the survey.",
    "Measurements were taken with a steel tape and checked against the drawings of the last refurbishment.",
    "The weather was dry and mild throughout, and the river level stayed close to its summer average.",
    "Where a defect could not be reached, its extent was estimated from photographs taken with a long lens.",
    "The previous survey recorded the same defect, and its extent has changed little since then.",
    "Pedestrian traffic was kept open during the inspection, with a marshal at each end of the span.",
    "No evidence of impact damage from river craft was found on the piers or the fenders.",
    "The paint system applied during the last refurbishment is generally performing as expected.",
    "Several joints are partly hidden by planting on the approaches, which should be cut back before the next visit.",
    "Readings from the tilt sensors installed on the piers were downloaded and show no trend over the year.",
    "The timber elements were tested with a moisture meter at regular intervals along each span.",
    "Drainage performance was checked by pouring water at the high points of the deck and following its path.",
]


def sentence_about(element=None):
    element = element or R.choice(ELEMENTS)
    loc = R.choice(LOCATIONS)
    obs = R.choice(OBS).format(n=R.randint(2, 40))
    rating = R.randint(1, 5)
    action = R.choice(ACTIONS)
    return (f"The {element} at {loc} showed {obs}. It was rated {rating}, {RATING_WORDS[rating]}, "
            f"and the proposed action is to {action if action != 'no action' else 'take no action for now'}.")


def body_paragraph(element=None, n=None):
    parts = []
    for _ in range(n or R.randint(3, 5)):
        parts.append(sentence_about(element) if R.random() < 0.55 else R.choice(FILLER))
    return " ".join(parts)


# ---------------------------------------------------------------- tables
BORDER = ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
          '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
          '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
          '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders>')


def cell(text, width, fill=None, header=False, right=False, span=1):
    tcpr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span > 1:
        tcpr += f'<w:gridSpan w:val="{span}"/>'
    if fill:
        tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    rpr = '<w:b/><w:color w:val="FFFFFF"/>' if header else ""
    ppr = '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>' + ('<w:jc w:val="right"/>' if right else "")
    inner = para(run(text, rpr + '<w:sz w:val="19"/>', toggles=not header), None, ppr)
    return f"<w:tc><w:tcPr>{tcpr}</w:tcPr>{inner}</w:tc>"


def row(cells, header=False):
    trpr = "<w:trPr><w:tblHeader/></w:trPr>" if header else '<w:trPr><w:cantSplit w:val="0"/></w:trPr>'
    return (f'<w:tr w:rsidR="{rsid()}" w:rsidTr="{rsid()}" w14:paraId="{para_id()}" w14:textId="77777777">'
            f"{trpr}{''.join(cells)}</w:tr>")


def table(widths, rows):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>{BORDER}'
            '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="85" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
            f"</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl>")


def findings_table():
    w = [700, 1500, 1650, 3510, 800, 1200]
    rows = [row([cell(h, x, "7B1E3A", header=True) for h, x in zip(("Ref", "Element", "Location", "Observation", "Rating", "Action"), w)], header=True)]
    for k in range(1, 57):
        rating = R.randint(1, 5)
        obs = R.choice(OBS).format(n=R.randint(2, 40))
        vals = (f"F{k:02d}", R.choice(ELEMENTS).capitalize(), R.choice(LOCATIONS).capitalize(), obs[0].upper() + obs[1:],
                str(rating), R.choice(ACTIONS).capitalize())
        fill = "F4EFE6" if k % 2 == 0 else None
        rows.append(row([cell(v, x, fill, right=(i == 4)) for i, (v, x) in enumerate(zip(vals, w))]))
    return table(w, rows)


def cost_table():
    w = [4160, 1200, 1000, 1400, 1600]
    items = [("Clean and reseal the expansion joints", 4, "no.", 850), ("Repaint the parapet base plates", 38, "m", 64),
             ("Replace the split decking boards", 22, "m", 145), ("Repair the spalled concrete at pier 2", 3, "m2", 620),
             ("Clear and test the drainage outlets", 12, "no.", 95), ("Replace the damaged lamp covers", 5, "no.", 180)]
    rows = [row([cell(h, x, "3E4A59", header=True) for h, x in zip(("Item", "Quantity", "Unit", "Rate (EUR)", "Amount (EUR)"), w)], header=True)]
    total = 0
    for name, q, u, rate in items:
        total += q * rate
        rows.append(row([cell(name, w[0]), cell(str(q), w[1], right=True), cell(u, w[2]), cell(f"{rate:,}", w[3], right=True),
                         cell(f"{q * rate:,}", w[4], right=True)]))
    rows.append(row([cell("Total, excluding access and contingency", sum(w[:4]), "F4EFE6", span=4), cell(f"{total:,}", w[4], "F4EFE6", right=True)]))
    return table(w, rows)


def scale_table():
    w = [1400, 7960]
    rows = [row([cell("Rating", w[0], "5C7F67", header=True), cell("Meaning", w[1], "5C7F67", header=True)], header=True)]
    for k, text in ((1, "Very good: no defect, or superficial defects only."),
                    (2, "Good: minor defects with no effect on the function of the element."),
                    (3, "Fair: defects that need attention at the next maintenance visit."),
                    (4, "Poor: defects that affect the function of the element and need repair within a year."),
                    (5, "Very poor: the element no longer performs its function; urgent action is needed.")):
        rows.append(row([cell(str(k), w[0], right=True), cell(text, w[1])]))
    return table(w, rows)


# ---------------------------------------------------------------- content controls
def sdt_inline(run_xml, tag, sid):
    return f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:id w:val="{sid}"/></w:sdtPr><w:sdtContent>{run_xml}</w:sdtContent></w:sdt>'


def sdt_block(paragraph_xml, tag, sid):
    return f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:id w:val="{sid}"/></w:sdtPr><w:sdtContent>{paragraph_xml}</w:sdtContent></w:sdt>'


# ---------------------------------------------------------------- the document
HEADINGS = []   # (level, text), in order, for the cached table of contents
body = []


def H(level, text, ppr=""):
    HEADINGS.append((level, text))
    body.append(heading(level, text, ppr))


def P(text, **kw):
    body.append(text_para(text, **kw))


# cover
body.append(para(run("Riverton Footbridge", toggles=False), "Title"))
body.append(para(run("Principal inspection and condition survey, 2025", toggles=False), "Subtitle"))
body.append(para(run(inner=picture(img_cover(), 360, 216, "Drawing of an arched footbridge")), None,
                 '<w:spacing w:before="480" w:after="480"/><w:jc w:val="center"/>'))
for line in ("Issue B, for review", "Survey carried out from 3 to 14 March 2025", "Report reference RFB-PI-2025-02"):
    body.append(para(run(line, '<w:color w:val="3E4A59"/><w:sz w:val="24"/>'), None, '<w:spacing w:after="60"/><w:jc w:val="center"/>'))
body.append(page_break())
TOC_AT = len(body)
body.append(None)       # the table of contents goes here once the headings are known
body.append(page_break())

H(1, "Summary")
P("This report records the principal inspection of the Riverton Footbridge, a three-span steel and timber "
  "structure carrying a shared path across the river between the market square and the park. The inspection "
  "covered every element that can be reached from the deck, the banks and a boat, and it follows the method "
  "described in section 2.")
body.append(para(run("The overall condition index of the structure is ") + sdt_inline(run("74 out of 100"), "goog_rdk_0", "-1736158329")
                 + run(", three points lower than at the previous survey. The decline comes mainly from the drainage and "
                       "the joints, which let water reach the bearings."), None, '<w:keepNext w:val="0"/>'))
P("The main findings are:")
for b in ("the expansion joints at both abutments have lost most of their sealant;",
          "two drainage outlets on span 2 are blocked, and water ponds on the deck after rain;",
          "the bearing at pier 2, east side, has moved by about 12 mm since it was last measured;",
          "the parapet base plates show paint breakdown and early corrosion over most of their length;",
          "the timber decking is sound apart from a group of split boards near the west abutment."):
    body.append(bullet(b))
P("None of these defects affects the safety of path users today. The recommendations in section 6 are ordered "
  "by urgency, and the cost estimate in section 7 covers the work proposed for the next twelve months.")

H(1, "Scope and method")
P(body_paragraph(n=4))
body.append(para(run("The inspection followed the guidance published at ") + hyperlink(add_rel("hyperlink", "https://example.org/footbridge-inspection-guide", True), "example.org/footbridge-inspection-guide")
                 + run(" and the condition rating scale given in appendix A. Photographs were catalogued with the naming rules of ")
                 + hyperlink(add_rel("hyperlink", "https://example.org/photo-records", True), "example.org/photo-records") + run(".")))
H(2, "Access arrangements")
P(body_paragraph(n=3))
body.append(sdt_block(text_para("Access to the river was agreed with the harbour office, which provided the boat and a pilot for two mornings."),
                      "goog_rdk_1", "1209466531"))
H(2, "Sequence of work")
for b in ("walk-over of the deck, parapets and approaches from both ends;",
          "close examination of the soffit, bearings and piers from the boat;",
          "measurement of bearing positions and joint gaps against the reference marks;",
          "drainage test at the high points of each span;",
          "review of the photographs and of the previous report, then rating of every element."):
    body.append(bullet(b, 2))
P(body_paragraph(n=3))

# The words before Figure 1 are counted so that the figure starts about a third of the way down its page:
# there, a line holding the picture at its own height leaves room for the caption, and the same line
# multiplied by the proportional spacing (276/240) does not.
LEAD = ("The footbridge was built in 1987 and refurbished in 2014. It has three spans of 18, 32 and 18 m on two "
        "reinforced concrete piers founded on piles, with a steel box girder under a timber deck. The deck is "
        "made of hardwood boards fixed to steel cross members at 600 mm centres, with an anti-slip surfacing on "
        "the ramps at both ends. The parapets are steel posts with a timber handrail and stainless steel mesh "
        "infill panels, and the lighting is carried on six columns bolted to the outer edge of the girder. "
        "Water drains through gaps between the boards onto the top flange of the girder, from where it runs to "
        "outlets at the low points of each span. The girder sits on elastomeric bearings at the piers and on "
        "guided sliding bearings at the abutments, which take the thermal movement of the whole structure. "
        "Both abutments are mass concrete on spread footings, with short wing walls retaining the approach "
        "ramps. The piers are founded on groups of six bored piles taken down into the gravel below the river "
        "bed, and each pier carries a timber fender on its upstream face to protect it from floating debris.")
LEAD_WORDS = int(os.environ.get("FIG_LEAD_WORDS", "178"))
H(1, "Structure description", '<w:pageBreakBefore/>')
P(" ".join(LEAD.split()[:LEAD_WORDS]).rstrip(",;.") + ".")
body.append(para(run(inner=picture(img_section(), 290, 414, "Schematic cross-section of a pier")), None,
                 '<w:keepNext w:val="0"/><w:spacing w:before="120" w:after="0"/><w:jc w:val="center"/>'))
body.append(para(run("Figure 1. Cross-section through pier 2, looking downstream (schematic, not to scale)."), "Caption"))
P(body_paragraph(n=4))
body.append(para(run(inner=picture(img_plan(), 468, 187, "Location plan")), None, '<w:jc w:val="center"/>'))
body.append(para(run("Figure 2. Location plan of the footbridge and its approaches."), "Caption"))
P(body_paragraph(n=3))

H(1, "Inspection findings")
P("Every defect recorded during the survey is listed below, with its location, its rating on the scale of "
  "appendix A and the proposed action. Photographs are referenced by the same code in the photo record.")
body.append(findings_table())
P("The ratings above were reviewed against the photographs after the survey and two were raised by one point.")
P(body_paragraph(n=4))

H(1, "Defects by element")
for el, sub in (("Deck and surfacing", ["Timber boards", "Anti-slip surfacing"]), ("Parapets and handrails", []),
                ("Bearings", ["Pier bearings", "Abutment bearings"]), ("Abutments and wing walls", []),
                ("Drainage", []), ("Lighting", [])):
    H(2, el)
    key = el.split()[0].lower()
    P(body_paragraph(n=R.randint(4, 6)))
    for s in sub:
        H(3, s)
        P(body_paragraph(n=R.randint(3, 5)))
    P(body_paragraph(n=R.randint(3, 5)))

H(1, "Recommendations")
P("The actions below are listed from the most to the least urgent. Items 1 to 4 should be done within twelve months.")
for b in ("Reseal the expansion joints at both abutments and clear the joint troughs.",
          "Clear the two blocked outlets on span 2 and add a grating to each.",
          "Survey the bearing at pier 2 again in six months and compare with the reference marks.",
          "Prepare and repaint the parapet base plates.",
          "Replace the split decking boards near the west abutment.",
          "Repair the spalled concrete on the east face of pier 2.",
          "Cut back the planting over the joints on both approaches.",
          "Replace the damaged lamp covers and test the lighting circuits."):
    body.append(bullet(b, 3))
P(body_paragraph(n=3))

H(1, "Cost estimate")
P("The estimate below covers items 1 to 6 of the recommendations, at the rates of the current maintenance framework.")
body.append(cost_table())
P("Access equipment, traffic management and a contingency of fifteen per cent should be added once the programme is agreed.")

H(1, "Appendix A. Condition rating scale")
body.append(scale_table())
P(body_paragraph(n=2))

H(1, "Appendix B. References")
for k, (label, url) in enumerate((("Footbridge inspection guide", "https://example.org/footbridge-inspection-guide"),
                                   ("Photo record naming rules", "https://example.org/photo-records"),
                                   ("Previous principal inspection report", "https://example.org/reports/2019"),
                                   ("Maintenance framework rates", "https://example.org/framework-rates"))):
    body.append(para(run(f"[{k + 1}] {label}: ") + hyperlink(add_rel("hyperlink", url, True), url), "ListParagraph"))


# ---------------------------------------------------------------- the table of contents in a content control
def toc_block():
    pages = TOC_PAGES or [3 + k for k in range(len(HEADINGS))]
    tab = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
    paras = [para(run("Contents", toggles=False), "TOCHeading")]
    for k, ((level, text), page) in enumerate(zip(HEADINGS, pages)):
        head = ""
        if k == 0:
            head = run(inner='<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'
                             '<w:fldChar w:fldCharType="separate"/>', toggles=False)
        entry = run(inner=t(text) + "<w:tab/>" + t(str(page)), toggles=False)
        tail = run(inner='<w:fldChar w:fldCharType="end"/>', toggles=False) if k == len(HEADINGS) - 1 else ""
        paras.append(para(head + entry + tail, f"TOC{level}", tab))
    return ('<w:sdt><w:sdtPr><w:id w:val="-1482736121"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
            '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr></w:sdtEndPr>'
            f"<w:sdtContent>{''.join(paras)}</w:sdtContent></w:sdt>")


body[TOC_AT] = toc_block()

# ---------------------------------------------------------------- stories, section
comments_rid = add_rel("comments", "comments.xml")
hdr_default = add_rel("header", "header1.xml")
hdr_first = add_rel("header", "header2.xml")
ftr_default = add_rel("footer", "footer1.xml")
ftr_first = add_rel("footer", "footer2.xml")
for kind, target in (("styles", "styles.xml"), ("settings", "settings.xml"), ("fontTable", "fontTable.xml"),
                     ("numbering", "numbering.xml"), ("theme", "theme/theme1.xml"), ("customXml", "../customXML/item1.xml")):
    add_rel(kind, target)

SECT = (f'<w:sectPr w:rsidR="{RSIDS[0]}" w:rsidSect="{RSIDS[2]}"><w:headerReference w:type="default" r:id="{hdr_default}"/>'
        f'<w:headerReference w:type="first" r:id="{hdr_first}"/><w:footerReference w:type="default" r:id="{ftr_default}"/>'
        f'<w:footerReference w:type="first" r:id="{ftr_first}"/>'
        '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>')
DOCUMENT = f"{HEAD}<w:document {NS}><w:body>{''.join(body)}{SECT}</w:body></w:document>"


def story(tag, paragraphs):
    return f"{HEAD}<w:{tag} {NS}>{paragraphs}</w:{tag}>"


HEADER1 = story("hdr", para(run("Riverton Footbridge, condition survey 2025", '<w:color w:val="7B1E3A"/><w:sz w:val="18"/>')
                             + run(rpr='<w:color w:val="3E4A59"/><w:sz w:val="18"/>', inner="<w:tab/>" + t("Issue B")),
                             "Header", '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="7B1E3A"/></w:pBdr>'))
HEADER2 = story("hdr", para("", "Header"))
FIELD = lambda instr: run(inner=f'<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve">{instr}</w:instrText>'
                                '<w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/>', rpr='<w:sz w:val="18"/>', toggles=False)
FOOTER1 = story("ftr", para(run("Page ", '<w:sz w:val="18"/>') + FIELD("PAGE") + run(" of ", '<w:sz w:val="18"/>') + FIELD("NUMPAGES"),
                            "Footer", '<w:jc w:val="center"/>'))
FOOTER2 = story("ftr", para(run("Synthetic test document: the names, places and figures are invented.", '<w:i/><w:sz w:val="16"/><w:color w:val="8A8D91"/>'),
                            "Footer", '<w:jc w:val="center"/>'))
COMMENTS = f"{HEAD}<w:comments {NS}></w:comments>"


# ---------------------------------------------------------------- styles
def pstyle(sid, name, ppr="", rpr="", based="Normal", nxt="Normal", extra=""):
    b = f'<w:basedOn w:val="{based}"/>' if based else ""
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>{b}<w:next w:val="{nxt}"/>{extra}'
            f"<w:qFormat/><w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>")


CAMBRIA = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:eastAsia="Cambria" w:cs="Cambria"/>'
STYLES = (f"{HEAD}<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>"
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/>'
          '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
          '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
          '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:semiHidden/></w:style>'
          + pstyle("Title", "Title", '<w:spacing w:before="2400" w:after="120" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>',
                   CAMBRIA + '<w:color w:val="7B1E3A"/><w:spacing w:val="-10"/><w:sz w:val="64"/><w:szCs w:val="64"/>')
          + pstyle("Subtitle", "Subtitle", '<w:spacing w:after="240"/><w:jc w:val="center"/>',
                   CAMBRIA + '<w:i/><w:color w:val="3E4A59"/><w:sz w:val="30"/><w:szCs w:val="30"/>')
          + pstyle("Heading1", "heading 1", '<w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="160" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="0"/>',
                   CAMBRIA + '<w:b/><w:bCs/><w:color w:val="7B1E3A"/><w:sz w:val="34"/><w:szCs w:val="34"/>')
          + pstyle("Heading2", "heading 2", '<w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="100" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="1"/>',
                   CAMBRIA + '<w:b/><w:bCs/><w:color w:val="3E4A59"/><w:sz w:val="26"/><w:szCs w:val="26"/>')
          + pstyle("Heading3", "heading 3", '<w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/>',
                   '<w:b/><w:i/><w:color w:val="A86B12"/>')
          + pstyle("Caption", "caption", '<w:spacing w:before="60" w:after="240" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>',
                   '<w:i/><w:iCs/><w:color w:val="3E4A59"/><w:sz w:val="18"/><w:szCs w:val="18"/>')
          + pstyle("ListParagraph", "List Paragraph", '<w:spacing w:after="80"/><w:ind w:left="720"/><w:contextualSpacing/>')
          + pstyle("TOCHeading", "TOC Heading", '<w:spacing w:before="240" w:after="240"/><w:outlineLvl w:val="9"/>',
                   CAMBRIA + '<w:b/><w:color w:val="7B1E3A"/><w:sz w:val="34"/>', based="Normal")
          + pstyle("TOC1", "toc 1", '<w:spacing w:before="120" w:after="60"/>', '<w:b/>')
          + pstyle("TOC2", "toc 2", '<w:spacing w:after="40"/><w:ind w:left="220"/>')
          + pstyle("TOC3", "toc 3", '<w:spacing w:after="40"/><w:ind w:left="440"/>', '<w:i/>')
          + pstyle("Header", "header", '<w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
          + pstyle("Footer", "footer", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
          + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>'
            '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="1F5C99"/><w:u w:val="single"/></w:rPr></w:style>'
          + "</w:styles>")


# ---------------------------------------------------------------- numbering, settings, fonts, theme
def abstract(aid, fmt, text, font=None):
    rf = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
    lvls = "".join(f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text.replace("%1", f"%{i + 1}")}"/>'
                   f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr>{rf}</w:lvl>' for i in range(3))
    return f'<w:abstractNum w:abstractNumId="{aid}"><w:multiLevelType w:val="hybridMultilevel"/>{lvls}</w:abstractNum>'


NUMBERING = (f"{HEAD}<w:numbering {NS}>" + abstract(0, "bullet", "•", "Calibri") + abstract(1, "lowerLetter", "%1)") + abstract(2, "decimal", "%1.")
             + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>')
SETTINGS = (f"{HEAD}<w:settings {NS}><w:zoom w:percent="+'"110"/><w:proofState w:spelling="clean" w:grammar="clean"/>'
            '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>'
            f'<w:rsids><w:rsidRoot w:val="{RSIDS[0]}"/>' + "".join(f'<w:rsid w:val="{r}"/>' for r in sorted(RSIDS)) + "</w:rsids>"
            '<w:themeFontLang w:val="en-GB"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" '
            'w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" '
            'w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>'
            '<w15:docId w15:val="{5B1A7C2E-0D4F-4E7A-9C31-2F6B8D0E4A17}"/></w:settings>')
FONTS = (f"{HEAD}<w:fonts {NS}>"
         '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
         '<w:font w:name="Cambria"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
         "</w:fonts>")


def style_fill(c):
    return f'<a:solidFill><a:schemeClr val="{c}"/></a:solidFill>'


THEME = (HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Riverton"><a:themeElements>'
         '<a:clrScheme name="Riverton"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="2B2D42"/></a:dk2><a:lt2><a:srgbClr val="F4EFE6"/></a:lt2><a:accent1><a:srgbClr val="7B1E3A"/></a:accent1>'
         '<a:accent2><a:srgbClr val="A86B12"/></a:accent2><a:accent3><a:srgbClr val="3E4A59"/></a:accent3><a:accent4><a:srgbClr val="5C7F67"/></a:accent4>'
         '<a:accent5><a:srgbClr val="C4A484"/></a:accent5><a:accent6><a:srgbClr val="8A8D91"/></a:accent6><a:hlink><a:srgbClr val="1F5C99"/></a:hlink>'
         '<a:folHlink><a:srgbClr val="6B4C9A"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Riverton"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Riverton"><a:fillStyleLst>' + style_fill("phClr") * 3 + '</a:fillStyleLst><a:lnStyleLst>'
         + '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' * 3 + '</a:lnStyleLst>'
         '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + '</a:effectStyleLst><a:bgFillStyleLst>'
         + style_fill("phClr") * 3 + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')
CUSTOM_ITEM = (HEAD + '<go:gDocsCustomXmlDataStorage xmlns:go="http://customooxmlschemas.google.com/" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" uri="GoogleDocsCustomDataVersion2">'
               '<go:docsCustomData xmlns:go="http://customooxmlschemas.google.com/" roundtripDataSignature="AMtx7mRivertonFixtureSignature0000000000=">'
               'CgMxLjA=</go:docsCustomData></go:gDocsCustomXmlDataStorage>')
CUSTOM_PROPS = (HEAD + '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" '
                'ds:itemID="{7E0F3C44-2B9A-4D61-8C5E-1A2B3C4D5E6F}"><ds:schemaRefs>'
                '<ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
                '<ds:schemaRef ds:uri="http://customooxmlschemas.google.com/"/></ds:schemaRefs></ds:datastoreItem>')
CORE = (HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Riverton Footbridge condition survey (synthetic test fixture)</dc:title><dc:creator>fixture generator</dc:creator>'
        '<cp:lastModifiedBy>fixture generator</cp:lastModifiedBy><cp:revision>3</cp:revision>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2025-03-20T09:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-03-27T16:30:00Z</dcterms:modified></cp:coreProperties>')

REL_TYPES = {"image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
             "hyperlink": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
             "comments": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
             "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
             "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
             "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
             "settings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
             "fontTable": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
             "numbering": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
             "theme": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
             "customXml": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"}
DOC_RELS = (HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="{rid}" Type="{REL_TYPES[k]}" Target="{escape(tg)}"' + (' TargetMode="External"' if ext else "") + "/>"
                      for rid, k, tg, ext in RELS) + "</Relationships>")
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
CONTENT_TYPES = (HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
                 f'<Override PartName="/word/document.xml" ContentType="{WML}.document.main+xml"/>'
                 f'<Override PartName="/word/styles.xml" ContentType="{WML}.styles+xml"/>'
                 f'<Override PartName="/word/settings.xml" ContentType="{WML}.settings+xml"/>'
                 f'<Override PartName="/word/fontTable.xml" ContentType="{WML}.fontTable+xml"/>'
                 f'<Override PartName="/word/numbering.xml" ContentType="{WML}.numbering+xml"/>'
                 f'<Override PartName="/word/comments.xml" ContentType="{WML}.comments+xml"/>'
                 f'<Override PartName="/word/header1.xml" ContentType="{WML}.header+xml"/>'
                 f'<Override PartName="/word/header2.xml" ContentType="{WML}.header+xml"/>'
                 f'<Override PartName="/word/footer1.xml" ContentType="{WML}.footer+xml"/>'
                 f'<Override PartName="/word/footer2.xml" ContentType="{WML}.footer+xml"/>'
                 '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                 '<Override PartName="/customXML/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
                 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                 "</Types>")
PKG_RELS = (HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            "</Relationships>")
CUSTOM_RELS = (HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>'
               "</Relationships>")

FIXED = (2025, 3, 27, 16, 30, 0)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    def put(name, data):
        z.writestr(zipfile.ZipInfo(name, FIXED), data, zipfile.ZIP_DEFLATED)
    put("[Content_Types].xml", CONTENT_TYPES)
    put("_rels/.rels", PKG_RELS)
    put("docProps/core.xml", CORE)
    put("word/document.xml", DOCUMENT)
    put("word/_rels/document.xml.rels", DOC_RELS)
    put("word/styles.xml", STYLES)
    put("word/settings.xml", SETTINGS)
    put("word/fontTable.xml", FONTS)
    put("word/numbering.xml", NUMBERING)
    put("word/comments.xml", COMMENTS)
    put("word/header1.xml", HEADER1)
    put("word/header2.xml", HEADER2)
    put("word/footer1.xml", FOOTER1)
    put("word/footer2.xml", FOOTER2)
    put("word/theme/theme1.xml", THEME)
    for name, data in MEDIA.items():
        put(f"word/media/{name}", data)
    put("customXML/item1.xml", CUSTOM_ITEM)
    put("customXML/itemProps1.xml", CUSTOM_PROPS)
    put("customXML/_rels/item1.xml.rels", CUSTOM_RELS)
print(f"{OUT}: {len(HEADINGS)} headings, {len(MEDIA)} pictures, {sum(1 for r in RELS if r[1] == 'hyperlink')} hyperlinks")
