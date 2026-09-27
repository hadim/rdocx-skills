"""Generate fixture-deck.pptx: a synthetic seven-slide board deck, written from scratch (invented text,
drawn pictures), 4:3, with a custom theme (Cambria titles, Calibri body, burgundy and ochre accents).

Usage: make_fixture_pptx.py [OUT]      (default fixture-deck.pptx; needs python-pptx and Pillow)

What the deck carries, for the acceptance workflow:
- title, content and title-only layouts, with placeholders and free text boxes on every slide
- groups of shapes with text (cards on slide 1, a timeline on slide 4), nested one level
- connectors glued to shapes at both ends, with arrowheads
- pictures on four slides; a table; a per-slide solid background on the title slide
- speaker notes on every slide
- explicit run fonts on some runs and inherited theme fonts on others
- one text box that overflows its frame and one that is filled to its last line
"""
import datetime, io, re, sys, zipfile
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

OUT = sys.argv[1] if len(sys.argv) > 1 else "fixture-deck.pptx"
BURGUNDY, OCHRE, SLATE, SAGE, CREAM, INK = (RGBColor(0x7B, 0x1E, 0x3A), RGBColor(0xA8, 0x6B, 0x12), RGBColor(0x3E, 0x4A, 0x59),
                                            RGBColor(0x5C, 0x7F, 0x67), RGBColor(0xF4, 0xEF, 0xE6), RGBColor(0x2B, 0x2D, 0x42))

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
L_TITLE, L_CONTENT, L_TITLE_ONLY = prs.slide_layouts[0], prs.slide_layouts[1], prs.slide_layouts[5]


def png(im):
    buf = io.BytesIO(); im.save(buf, "PNG"); buf.seek(0); return buf


def img_arch(w=800, h=420):
    im = Image.new("RGB", (w, h), (244, 239, 230)); d = ImageDraw.Draw(im)
    d.rectangle([0, int(h * 0.72), w, h], fill=(96, 128, 150))
    d.arc([int(w * 0.1), int(h * 0.2), int(w * 0.9), int(h * 1.25)], 180, 360, fill=(123, 30, 58), width=12)
    d.line([20, int(h * 0.72), w - 20, int(h * 0.72)], fill=(62, 74, 89), width=9)
    return im


def img_plan(w=900, h=560):
    im = Image.new("RGB", (w, h), (250, 250, 247)); d = ImageDraw.Draw(im)
    d.polygon([(0, 220), (w, 120), (w, 330), (0, 430)], fill=(170, 196, 214))
    d.line([330, 60, 450, 500], fill=(123, 30, 58), width=12)
    for x in range(30, w, 90):
        d.ellipse([x, 470, x + 30, 500], fill=(92, 127, 103))
    return im


def img_detail(w=600, h=600, seed=0):
    im = Image.new("RGB", (w, h), (225, 222, 215)); d = ImageDraw.Draw(im)
    for k in range(0, w, 60):
        d.rectangle([k, 0, k + 52, h], fill=(196 - seed * 9, 170 - seed * 6, 132))           # boards
    d.line([0, h // 2 + seed * 20, w, h // 2 - 40], fill=(90, 70, 50), width=4)              # a crack
    return im


def text_box(slide, x, y, w, h, paragraphs, size=14, color=INK, font=None, bold=False, align=None, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = wrap; tf.auto_size = MSO_AUTO_SIZE.NONE
    for k, text in enumerate(paragraphs):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.color.rgb = color; r.font.bold = bold or None
        if font:
            r.font.name = font
        if align:
            p.alignment = align
    return tb


def title(slide, text):
    slide.shapes.title.text = text
    r = slide.shapes.title.text_frame.paragraphs[0].runs[0]
    r.font.color.rgb = BURGUNDY; r.font.size = Pt(30)
    slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    return slide.shapes.title


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def card(group_parent, x, y, w, h, figure, label, fill):
    g = group_parent.add_group_shape()
    box = g.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    box.fill.solid(); box.fill.fore_color.rgb = fill; box.line.fill.background()
    box.text_frame.text = ""
    t1 = g.shapes.add_textbox(x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), Inches(0.7))
    t1.text_frame.text = figure
    r = t1.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    r.font.name = "Cambria"
    t2 = g.shapes.add_textbox(x + Inches(0.1), y + Inches(0.8), w - Inches(0.2), h - Inches(0.9))
    t2.text_frame.word_wrap = True; t2.text_frame.text = label
    r = t2.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(12); r.font.color.rgb = RGBColor(255, 255, 255)
    return g, box


def glue(slide, a, b, a_site=3, b_site=1):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
    c.begin_connect(a, a_site); c.end_connect(b, b_site)
    c.line.color.rgb = SLATE; c.line.width = Pt(1.5)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return c


# 0. title slide
s = prs.slides.add_slide(L_TITLE)
s.background.fill.solid(); s.background.fill.fore_color.rgb = BURGUNDY
s.shapes.title.text = "Riverton Footbridge"
for r in s.shapes.title.text_frame.paragraphs[0].runs:
    r.font.color.rgb = RGBColor(255, 255, 255); r.font.size = Pt(44)
s.placeholders[1].text = "Refurbishment options for the board, 2 April 2025"
for r in s.placeholders[1].text_frame.paragraphs[0].runs:
    r.font.color.rgb = CREAM; r.font.size = Pt(20); r.font.italic = True
for ph, top, height in ((s.shapes.title, 1.1, 1.3), (s.placeholders[1], 2.4, 0.8)):
    ph.left, ph.top, ph.width, ph.height = Inches(0.75), Inches(top), Inches(8.5), Inches(height)
s.shapes.add_picture(png(img_arch()), Inches(2.5), Inches(3.6), width=Inches(5))
text_box(s, Inches(0.5), Inches(6.6), Inches(9), Inches(0.4), ["Synthetic test deck: the names, places and figures are invented."],
         size=10, color=CREAM, align=PP_ALIGN.CENTER)
notes(s, "Welcome the board. The survey was carried out in March; this deck summarises it and asks for a decision on the option.")

# 1. condition at a glance: grouped cards, glued connectors, a picture
s = prs.slides.add_slide(L_TITLE_ONLY)
title(s, "Condition at a glance")
text_box(s, Inches(0.5), Inches(1.35), Inches(9), Inches(0.5),
         ["The structure is safe to use; the decline since the last survey comes from water getting into the joints."],
         size=14, font="Calibri")
outer = s.shapes.add_group_shape()
boxes = []
for k, (fig, label, fill) in enumerate((("74", "Condition index out of 100, down 3 since the last survey", BURGUNDY),
                                        ("56", "Defects recorded, 9 of them rated poor or very poor", SLATE),
                                        ("12 mm", "Movement of the bearing at pier 2 since the reference survey", OCHRE))):
    g, box = card(outer.shapes, Inches(0.5 + 3.1 * k), Inches(2.0), Inches(2.8), Inches(1.9), fig, label, fill)
    boxes.append(box)
for a, b in zip(boxes, boxes[1:]):
    glue(s, a, b)
s.shapes.add_picture(png(img_plan()), Inches(0.5), Inches(4.2), height=Inches(2.8))
text_box(s, Inches(5.2), Inches(4.3), Inches(4.3), Inches(2.6),
         ["The footbridge carries about 3,000 people a day between the market square and the park.",
          "Closing it for works means a detour of 1.4 km by the road bridge upstream."], size=13)
notes(s, "The three figures come from the survey report. The bearing movement is the one item to watch closely.")

# 2. main defects: body placeholder bullets, pictures, captions
s = prs.slides.add_slide(L_CONTENT)
title(s, "Main defects")
body = s.placeholders[1]
body.left, body.top, body.width, body.height = Inches(0.5), Inches(1.4), Inches(5.2), Inches(5.4)
tf = body.text_frame
for k, line in enumerate(("Expansion joints at both ends have lost most of their sealant",
                          "Two drainage outlets on the main span are blocked",
                          "The pier 2 bearing has moved by about 12 mm",
                          "Paint breakdown at the parapet base plates",
                          "Split decking boards near the west end")):
    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
    p.text = line
    p.runs[0].font.size = Pt(18)
for k in range(2):
    s.shapes.add_picture(png(img_detail(seed=k)), Inches(6.0), Inches(1.4 + 2.75 * k), width=Inches(2.3))
    text_box(s, Inches(6.0), Inches(3.75 + 2.75 * k), Inches(3.5), Inches(0.35), [("Split boards, span 1", "Joint gap, east end")[k]],
             size=10, color=SLATE)
notes(s, "Photographs are from the survey record; the full list of 56 defects is in the report.")

# 3. options: a table and text boxes
s = prs.slides.add_slide(L_TITLE_ONLY)
title(s, "Three options")
rows = (("", "A. Repair", "B. Refurbish", "C. Replace deck"),
        ("Cost (EUR)", "46,000", "212,000", "690,000"),
        ("Closure", "2 weeks", "7 weeks", "5 months"),
        ("Next major work", "2031", "2040", "2055"))
gf = s.shapes.add_table(4, 4, Inches(0.5), Inches(1.5), Inches(9), Inches(2.4))
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        c = gf.table.cell(i, j); c.text = v
        for r in c.text_frame.paragraphs[0].runs:
            r.font.size = Pt(15)
        if i == 0:
            c.fill.solid(); c.fill.fore_color.rgb = SLATE
text_box(s, Inches(0.5), Inches(4.3), Inches(4.3), Inches(2.4),
         ["Option B is recommended: it deals with every defect rated poor or worse and keeps the path open "
          "for all but seven weeks, outside the summer season."], size=15)
text_box(s, Inches(5.2), Inches(4.3), Inches(4.3), Inches(2.4),
         ["Option A leaves the bearing and the drainage as they are and brings the next closure forward.",
          "Option C is only worth it if the timber deck is replaced by a steel one."], size=13, color=SLATE)
notes(s, "Costs exclude access and contingency. The dates for the next major work assume the current rate of decline.")

# 4. programme: a grouped timeline of chevrons, glued connectors, an overflowing box
s = prs.slides.add_slide(L_TITLE_ONLY)
title(s, "Programme for option B")
tl = s.shapes.add_group_shape()
chevrons = []
for k, (label, fill) in enumerate((("Design", SLATE), ("Tender", SAGE), ("Works", BURGUNDY), ("Handover", OCHRE))):
    ch = tl.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0.5 + 2.25 * k), Inches(1.6), Inches(2.2), Inches(0.9))
    ch.fill.solid(); ch.fill.fore_color.rgb = fill; ch.line.fill.background()
    ch.text_frame.text = label
    r = ch.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    chevrons.append(ch)
milestones = []
for k, (when, what) in enumerate((("May 2025", "Survey of the bearing"), ("Sep 2025", "Contract award"),
                                  ("Jan 2026", "Path closed"), ("Mar 2026", "Path reopened"))):
    m = s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(1.3 + 2.25 * k), Inches(3.1), Inches(0.5), Inches(0.5))
    m.fill.solid(); m.fill.fore_color.rgb = CREAM; m.line.color.rgb = SLATE
    milestones.append(m)
    text_box(s, Inches(0.8 + 2.25 * k), Inches(3.7), Inches(1.6), Inches(0.8), [when, what], size=11, align=PP_ALIGN.CENTER)
for a, b in zip(milestones, milestones[1:]):
    glue(s, a, b)
text_box(s, Inches(0.5), Inches(5.0), Inches(4.2), Inches(1.2),
         ["Risks: the tender may attract few bids in the autumn; the bearing survey may show that a jacking operation "
          "is needed, which would add three weeks; the river authority must approve the scaffold in the channel before "
          "the works can start, and its approval has taken up to ten weeks on similar schemes; a wet winter would slow "
          "the painting."], size=13)
text_box(s, Inches(5.2), Inches(5.0), Inches(4.3), Inches(1.5),
         ["The closure is planned for the quietest months of the year.", "A shuttle bus will run during the works."], size=13)
notes(s, "The risk box is deliberately dense: it is the one the board asked to see in full.")

# 5. costs and funding: text filled to its last line, a small picture
s = prs.slides.add_slide(L_TITLE_ONLY)
title(s, "Costs and funding")
text_box(s, Inches(0.5), Inches(1.5), Inches(5.6), Inches(2.25),
         ["Works: EUR 184,000", "Access and traffic management: EUR 16,000", "Contingency at 15 per cent: EUR 30,000",
          "Total: EUR 230,000"], size=18)
text_box(s, Inches(0.5), Inches(4.2), Inches(5.6), Inches(2.4),
         ["Funding: the maintenance reserve covers EUR 150,000; the rest is asked from the regional fund for "
          "active travel, which opens its next round in June."], size=15, color=SLATE)
s.shapes.add_picture(png(img_detail(seed=2)), Inches(6.6), Inches(1.5), width=Inches(2.9))
notes(s, "Figures rounded to the nearest thousand.")

# 6. decision
s = prs.slides.add_slide(L_CONTENT)
title(s, "Decision requested")
tf = s.placeholders[1].text_frame
for k, line in enumerate(("Approve option B and its budget of EUR 230,000",
                          "Approve the application to the regional fund",
                          "Note the programme and the closure from January to March 2026")):
    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
    p.text = line
notes(s, "If the board prefers option A, the bearing survey still goes ahead in May.")

prs.core_properties.title = "Riverton Footbridge refurbishment options (synthetic test fixture)"
prs.core_properties.created = datetime.datetime(2025, 3, 28, 9, 0, 0)
prs.core_properties.modified = datetime.datetime(2025, 3, 31, 18, 0, 0)
prs.core_properties.revision = 4
prs.core_properties.author = "fixture generator"
prs.core_properties.last_modified_by = "fixture generator"
prs.core_properties.subject = ""
prs.core_properties.comments = ""
prs.core_properties.keywords = ""
buf = io.BytesIO(); prs.save(buf)

# custom theme: colours and fonts of the fixture, in place of the template's
CLR = ('<a:clrScheme name="Riverton"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
       '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="2B2D42"/></a:dk2>'
       '<a:lt2><a:srgbClr val="F4EFE6"/></a:lt2><a:accent1><a:srgbClr val="7B1E3A"/></a:accent1>'
       '<a:accent2><a:srgbClr val="A86B12"/></a:accent2><a:accent3><a:srgbClr val="3E4A59"/></a:accent3>'
       '<a:accent4><a:srgbClr val="5C7F67"/></a:accent4><a:accent5><a:srgbClr val="C4A484"/></a:accent5>'
       '<a:accent6><a:srgbClr val="8A8D91"/></a:accent6><a:hlink><a:srgbClr val="1F5C99"/></a:hlink>'
       '<a:folHlink><a:srgbClr val="6B4C9A"/></a:folHlink></a:clrScheme>')
FONT = ('<a:fontScheme name="Riverton"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>')
zin = zipfile.ZipFile(buf)
FIXED = (2025, 3, 31, 18, 0, 0)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if re.fullmatch(r"ppt/theme/theme\d+\.xml", info.filename):
            x = data.decode()
            x = re.sub(r"<a:clrScheme .*?</a:clrScheme>", CLR, x, flags=re.S)
            x = re.sub(r"<a:fontScheme .*?</a:fontScheme>", FONT, x, flags=re.S)
            x = re.sub(r'(<a:theme [^>]*name=")[^"]*"', r'\1Riverton"', x)
            data = x.encode()
        z.writestr(zipfile.ZipInfo(info.filename, FIXED), data, zipfile.ZIP_DEFLATED)
print(f"{OUT}: {len(prs.slides)} slides")
