# pptx recipes (each block is run by the test suite on the pinned build)

Conventions, as set up in SKILL.md: `R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, `SKILL` is
this skill's folder, and Python blocks run with `PYTHONPATH="$SKILL/scripts" $R/python`, so that
`import pptx_ops` works. `deck.pptx` is the input; outputs go to new files. Blocks run in order in one
folder.

## Read the deck

```bash
$R/rpptx outline --notes deck.pptx
$R/rpptx text --json deck.pptx > deck.json
$R/python "$SKILL/scripts/pptx_ops.py" shapes deck.pptx --slide 2
```

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
for k, slide in enumerate(prs.slides):
    title = slide.shapes.title.text if slide.shapes.title else ""
    texts = [sh.text for _, sh in pptx_ops.walk(slide.shapes) if sh.has_text_frame and sh.text]
    print(k + 1, title, len(texts), (slide.notes_text or "")[:40])    # notes_text is None without notes
```

## Counted replacements, all or nothing

```bash
$R/rpptx replace deck.pptx -p "2 April 2025" -v "9 April 2025" --expect 1 -o dated.pptx
```

```python
import pptx_ops
pptx_ops.replace_batch("deck.pptx", "edited.pptx", [
    ("EUR 230,000", "EUR 236,000", 2),
    ("seven weeks", "eight weeks", 1),
])
```

## Replace in one slide or one frame

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
prs.slides[2].try_replace_text("12 mm", "14 mm", expect=1, notes=False)   # the deck holds it twice: slides 2 and 3
k = next(i for i, sh in enumerate(prs.slides[5].shapes) if sh.has_text_frame and "Total:" in sh.text)
prs.slides[5].shapes[k].text_frame.try_replace_text("230,000", "236,000", expect=1)  # this frame only, not slide 7
pptx_ops.save_atomic(prs, "scoped.pptx", "deck.pptx")
```

## Edit a run without losing its formatting

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
k = next(i for i, sh in enumerate(prs.slides[5].shapes) if sh.has_text_frame and sh.text.startswith("Works:"))
prs.slides[5].shapes[k].text_frame.paragraphs[0].runs[0].text = "Works: EUR 188,000"
run = prs.slides[5].shapes[k].text_frame.paragraphs[0].runs[0]      # re-fetch after Run.text
assert run.font.size is not None
pptx_ops.save_atomic(prs, "run-edited.pptx", "deck.pptx")
```

## Move, resize, restyle a shape; add a shape, a picture and a connector

```python
import io, pptx_ops, rpptx
from rpptx.dml.color import RGBColor
from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from rpptx.util import Inches, Pt
prs = rpptx.Presentation("deck.pptx")
s = prs.slides[3]
box = next(i for i, sh in enumerate(s.shapes) if sh.has_text_frame and sh.text.startswith("Option B is recommended"))
prs.slides[3].shapes[box].top = prs.slides[3].shapes[box].top + Inches(0.2)
badge = prs.slides[3].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8), Inches(0.4), Inches(1.5), Inches(0.5))
badge.fill.solid(); badge.fill.fore_color.rgb = RGBColor(0x7B, 0x1E, 0x3A); badge.line.width = Pt(0.75)
badge.text_frame.text = "Decision"
png = io.BytesIO(pptx_ops.solid_png(400, 200, (220, 220, 220)))     # or open("photo.png", "rb")
prs.slides[3].shapes.add_picture(png, Inches(7), Inches(5.2), width=Inches(2))
prs.slides[3].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7), Inches(1), Inches(8), Inches(1))
pptx_ops.save_atomic(prs, "shapes.pptx", "deck.pptx")
```

## Shadow, dashes and arrowheads, a connector without the theme effect, another preset

```python
import pptx_ops, rpptx
from rpptx.dml.color import RGBColor
from rpptx.enum.dml import MSO_ARROWHEAD_LENGTH, MSO_ARROWHEAD_STYLE, MSO_ARROWHEAD_WIDTH, MSO_LINE_DASH_STYLE
from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from rpptx.util import Inches, Pt
prs = rpptx.Presentation("deck.pptx")
card = prs.slides[4].shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(6.4), Inches(2.5), Inches(0.6))
card.fill.solid(); card.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)   # add_shape writes the theme style
card.line.color.rgb = RGBColor(0x80, 0x80, 0x80)                          # (accent1): direct values replace it
card.auto_shape_type = MSO_SHAPE.RECTANGLE                                # change the preset in place
card.shadow.visible = True; card.shadow.color.rgb = RGBColor(0, 0, 0); card.shadow.alpha = 0.35
card.shadow.blur_radius = Pt(4); card.shadow.distance = Pt(3); card.shadow.direction = 45.0; card.shadow.align = "tl"
arrow = prs.slides[4].shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3.2), Inches(6.7), Inches(5), Inches(6.7))
arrow.theme_effect_index = 0                     # add_connector references effect 1, a shadow in the default theme
arrow.line.color.rgb = RGBColor(0x40, 0x40, 0x40); arrow.line.width = Pt(1.5)
arrow.line.dash_style = MSO_LINE_DASH_STYLE.DASH
arrow.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE
arrow.line.tail_end.width = MSO_ARROWHEAD_WIDTH.WIDE; arrow.line.tail_end.length = MSO_ARROWHEAD_LENGTH.LONG
pptx_ops.save_atomic(prs, "styled.pptx", "deck.pptx")
```

## Grow a table, fill a group

```python
import pptx_ops, rpptx
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches
prs = rpptx.Presentation("deck.pptx")
t = next(k for k, sh in enumerate(prs.slides[3].shapes) if sh.has_table)
n = len(prs.slides[3].shapes[t].table.rows)
prs.slides[3].shapes[t].table.rows.add_row(n - 1)        # before the last row, formatted like its neighbour
prs.slides[3].shapes[t].table.cell(n - 1, 0).text = "Contingency"
g = len(prs.slides[3].shapes)
prs.slides[3].shapes.add_group_shape()
prs.slides[3].shapes[g].shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7), Inches(6), Inches(2), Inches(0.4))
prs.slides[3].shapes[g].shapes.add_textbox(Inches(7), Inches(6.5), Inches(2), Inches(0.4)).text_frame.text = "Legend"
pptx_ops.save_atomic(prs, "grown.pptx", "deck.pptx")
```

## Slides: notes, order, visibility, new slide

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
prs.slides[2].notes_text = "Mention the photo record for each defect."
prs.slides.move(5, 4)                         # costs before the programme
prs.slides[6].hidden = True
prs.slides.add_slide(prs.slide_layouts[1])
new = prs.slides[len(prs.slides) - 1]
new.shapes.title.text = "Questions"
pptx_ops.save_atomic(prs, "reordered.pptx", "deck.pptx")
```

## Import a slide from another deck

```python
import pptx_ops, rpptx
other = rpptx.Presentation()                     # stands for rpptx.Presentation("other.pptx")
other.slides.add_slide(other.slide_layouts[5])
other.slides[0].shapes.title.text = "Appendix: survey method"
prs = rpptx.Presentation("deck.pptx")
# without layout=, deck.pptx needs a layout named like the source's, else RpptxError "no layout named ..."
prs.slides.import_slide(other.slides[0], layout=prs.slide_layouts[5], index=1)   # second slide
pptx_ops.save_atomic(prs, "imported.pptx", "deck.pptx")
```

## Replace a picture, keeping its frame

```python
import io, pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
i = next(k for k, sh in enumerate(prs.slides[0].shapes) if int(sh.shape_type or 0) == 13)
w, h = pptx_ops.image_size(prs.slides[0].shapes[i].image.blob)
prs.slides[0].shapes[i].replace_image(io.BytesIO(pptx_ops.solid_png(w, h, (60, 100, 140))))   # same aspect ratio
pptx_ops.save_atomic(prs, "picture.pptx", "deck.pptx")
```

## One-shot slide edits, footers and metadata from the CLI

```bash
$R/rpptx slide add deck.pptx --layout "Title Only" --at 8 -o s1.pptx    # layout name or one-based number
$R/rpptx slide move s1.pptx 6 --to 5 -o s2.pptx
$R/rpptx notes set s2.pptx 8 --text "Questions from the board." -o s3.pptx
$R/rpptx footer s3.pptx --slide-number --footer "Riverton Footbridge" --skip-title -o s4.pptx
$R/rpptx meta set s4.pptx --title "Footbridge options" --author "Claude" -o s5.pptx
$R/rpptx meta get s5.pptx
$R/rpptx fit --json s5.pptx > fit.json || [ $? -eq 1 ]      # exit 1: a frame overflows, listed in fit.json
```

## Brand a deck: theme, logo, slide numbers, transitions, alt text

```python
import io, pptx_ops, rpptx
from rpptx.util import Inches
prs = rpptx.Presentation("deck.pptx")
theme = prs.slide_master.theme
theme.colors["accent1"] = "7B1E3A"                    # every shape that uses accent1 follows
theme.fonts.major.latin = "Georgia"                   # headings
logo = io.BytesIO(pptx_ops.solid_png(60, 60, (123, 30, 58)))
prs.slide_master.shapes.add_picture(logo, Inches(9.2), Inches(0.1), width=Inches(0.6))   # on every slide
prs.slide_master.shapes[len(prs.slide_master.shapes) - 1].alt_text = "Company logo"
prs.set_header_footer(footer="Riverton Footbridge")  # slide numbers on, title slides skipped
prs.slides[0].transition.type = "fade"
prs.slides[0].transition.apply_to_all()
prs.core_properties.title = "Footbridge options"
pptx_ops.save_atomic(prs, "branded.pptx", "deck.pptx")
```

## Bullets, numbering, a fitted text box, a glued connector

```python
import pptx_ops, rpptx
from rpptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from rpptx.util import Inches, Pt
prs = rpptx.Presentation("deck.pptx")
slide = prs.slides.add_slide(prs.slide_layouts.get_by_name("Title Only"))
slide.shapes.title.text = "Next steps"
box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4), Inches(1.5))
box.text_frame.text = "Confirm the bearing survey"
for text in ("Price option B", "Book the closure"):
    box.text_frame.add_paragraph().text = text
for p in box.text_frame.paragraphs:
    p.auto_number = "arabicPeriod"                    # 1. 2. 3.; p.bullet = "\u2022" for bullets
box.text_frame.fit_text(max_size=24)                  # the largest whole size that fits
a = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(2), Inches(1.5), Inches(0.8))
b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(4), Inches(1.5), Inches(0.8))
line = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, 0, 0, 0, 0)
line.begin_connect(a, 3); line.end_connect(b, 1)      # right side of a to left side of b, glued
a.alt_text, b.alt_text = "Survey", "Closure"
pptx_ops.save_atomic(prs, "steps.pptx", "deck.pptx")
```

## Raw XML when no API exists

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
k = next(i for i, sh in enumerate(prs.slides[3].shapes) if sh.has_text_frame and sh.text.startswith("Option B is recommended"))
xml = prs.slides[3].shapes[k].text_frame.xml          # bytes: the p:txBody
prs.slides[3].shapes[k].text_frame.replace_xml(xml.replace(b"<a:bodyPr", b'<a:bodyPr numCol="2"', 1))  # two columns
try:
    prs.slides[3].shapes[k].replace_xml(b"<p:sp/>")   # refused, deck unchanged
except ValueError as e:
    print("refused:", e)
pptx_ops.save_atomic(prs, "raw.pptx", "deck.pptx")
```

## Comments

```bash
$R/rpptx comment add deck.pptx --slide 4 --author "Reviewer" --text "Source for the costs?" --date 2026-09-27T12:00:00Z -o c1.pptx
ID=$($R/rpptx comment list --json c1.pptx | $R/python -c 'import json,sys; print(json.load(sys.stdin)["comments"][0]["id"])')
$R/rpptx comment reply c1.pptx --id "$ID" --author "Author" --text "Framework rates, 2025." --date 2026-09-27T12:05:00Z -o c2.pptx
$R/rpptx comment resolve c2.pptx --id "$ID" -o c3.pptx
$R/rpptx comment list c3.pptx
```

## Shape links, slide jumps, a comment on a shape

```python
import pptx_ops, rpptx
prs = rpptx.Presentation("deck.pptx")
k = next(i for i, sh in enumerate(prs.slides[3].shapes) if sh.has_text_frame and sh.text.startswith("Option B is recommended"))
prs.slides[3].shapes[k].click_action.target_slide = prs.slides[4]          # a click jumps to slide 5
assert prs.slides[3].shapes[k].click_action.target_slide == prs.slides[4]  # Slide handles compare with ==
prs.slides[0].shapes[3].click_action.hyperlink.address = "https://example.com/footbridge"   # a web link on a shape
author = "{6F1B3A52-0000-4000-8000-000000000001}"
prs.add_comment_author(id=author, name="Claude", user_id="Claude", provider_id="None", initials="C")
prs.slides[3].add_comment(id="{6F1B3A52-0000-4000-8000-000000000002}", author_id=author, created="2026-10-01T12:00:00Z",
                          text="Which option did the board pick?", shape_id=prs.slides[3].shapes[k].shape_id)
pptx_ops.save_atomic(prs, "links.pptx", "deck.pptx")
```

## Check the fit, then render

```bash
$R/python "$SKILL/scripts/pptx_ops.py" overflow deck.pptx
slides=$(mktemp -d)                           # render into a new folder each time
$R/rpptx render deck.pptx -o "$slides" --slide 5 --dpi 80
$R/rpptx convert deck.pptx --to pdf -o deck.pdf   # a new path: convert --to pdf overwrites
$R/rpptx validate deck.pptx
```

```python
import rpptx
prs = rpptx.Presentation("deck.pptx")
tight = [(f.slide_index + 1, f.name) for f in prs.text_layout(width_factor=0.95) if f.overflow]
print(tight)                                  # frames that would overflow in a 5 % narrower box
```

## A deck from a .potx template

```python
import pptx_ops, rpptx, zipfile
with zipfile.ZipFile("deck.pptx") as src, zipfile.ZipFile("deck.potx", "w") as dst:   # a .potx for this example
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename == "[Content_Types].xml":
            data = data.replace(b"presentation.main+xml", b"template.main+xml")
        dst.writestr(info, data)
prs = rpptx.Presentation("deck.potx")
prs.slides[0].shapes.title.text = "From the template"
pptx_ops.save_atomic(prs, "from-template.pptx")            # saved as .pptx: a presentation
with zipfile.ZipFile("from-template.pptx") as z:
    assert b"presentationml.presentation.main+xml" in z.read("[Content_Types].xml")
```
