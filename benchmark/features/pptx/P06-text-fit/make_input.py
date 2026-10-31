"""P06 input: a 3-slide deck with 13 text frames of known fit, and truth.json. The master's autofit is off
(noAutofit) so that no renderer may shrink text and the truth does not depend on autofit. Each overflow is gross
(at least 1.5 times the box) and each fitting frame keeps at least 13 % slack, so that the truth holds for
PowerPoint, LibreOffice and rpptx alike despite their small layout differences.

Text boxes have no autofit either (python-pptx's spAutoFit default is removed): a box keeps its size.

Overflow: a wrapped 2-inch text box with a long sentence; an unwrapped line wider than its box; a title
placeholder with a long title (geometry and 44 pt inherited, centred: it spills above and below); a body
placeholder with 10 bullets (geometry and 32 pt inherited); a table whose long cell grows its row until the table
runs past the slide's bottom edge. Fit: two titles, a roomy box, a near miss (3 short lines, 13 % slack), a badge
shape, a small table, a table whose cell is taller than its stored row height (PowerPoint, LibreOffice and
rpptx's renderer grow the row; the table stays on the slide, clear of other shapes), a footnote, a source line."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402
from inputs import Inches, MSO_SHAPE, NS, Pt  # noqa: E402
from pptx.enum.text import MSO_AUTO_SIZE  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
for body_pr in prs.slide_master.element.iter("{%s}bodyPr" % NS["a"]):
    for auto in body_pr.findall("a:normAutofit", NS):
        body_pr.remove(auto)
        body_pr.append(body_pr.makeelement("{%s}noAutofit" % NS["a"], {}))

s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Fit check: boxes"
inputs.textbox(s.shapes, 0.8, 1.8, 2.0, 0.6, ["Our customers in the northern region renewed their contracts earlier "
                                              "than planned, which lifted revenue and reduced the sales effort for "
                                              "the whole quarter."], size=18, name="Box narrow")
inputs.textbox(s.shapes, 3.6, 1.8, 5.0, 2.0, ["Revenue up 12 %"], size=18, name="Box roomy")
near = inputs.textbox(s.shapes, 9.2, 1.8, 3.4, 1.15, ["North: on plan"], size=18, name="Box near miss")
for line in ("South: behind", "West: ahead"):
    p = near.text_frame.add_paragraph()
    inputs.runs(p, [(line, {"size": 18})])
inputs.textbox(s.shapes, 0.8, 4.6, 2.5, 0.5, ["Quarterly revenue grew in every region this year"], size=20,
               name="Unwrapped line", wrap=False)
badge = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(4.6), Inches(2.0), Inches(0.6))
badge.name = "Badge"
badge.text_frame.text = "Q3"
badge.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

s = prs.slides.add_slide(inputs.layout(prs, "content"))
s.shapes.title.text = ("Northwind regional performance review covering revenue, margin, customer churn, delivery "
                       "times and hiring across all regions")
inputs.bullets(s.placeholders[1].text_frame, [(0, [f"Point {k}: regional figure to discuss"]) for k in range(1, 11)])

s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Fit check: tables"
inputs.table(s.shapes, 0.8, 1.7, 6.0, 0.8, [["Region", "Revenue"], ["North", "1,240"]], name="Fit table")
big = inputs.table(s.shapes, 0.8, 5.0, 11.7, 1.2, [["Region", "Comment", "Owner"], ["North", "", "Dana"],
                                                    ["South", "On plan", "Sam"]], name="Overflow table")
big.table.columns[1].width = Inches(4.0)
big.table.columns[0].width, big.table.columns[2].width = Inches(2.0), Inches(5.7)
for row in big.table.rows:
    row.height = Inches(0.4)
cell = big.table.cell(1, 1)
cell.text = ("Renewals came early in the North, " * 9).strip()
cell.text_frame.paragraphs[0].runs[0].font.size = Pt(18)
inputs.textbox(s.shapes, 7.2, 1.7, 5.3, 0.4, ["Figures in EUR thousand"], size=12, name="Footnote")
inputs.textbox(s.shapes, 7.2, 2.3, 5.3, 0.4, ["Source: finance, September close"], size=12, name="Source")
grow = inputs.table(s.shapes, 0.8, 2.9, 11.7, 0.8, [["Risk", "Mitigation"], ["Hiring", ""]], name="Growing table")
grow.table.columns[0].width, grow.table.columns[1].width = Inches(2.0), Inches(9.7)
for row in grow.table.rows:
    row.height = Inches(0.4)
cell = grow.table.cell(1, 1)
cell.text = ("Start interviews in November with two agencies; " * 4).strip()
cell.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
for slide in prs.slides:  # no text box resizes itself
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder:
            shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
inputs.save(prs, out / "deck.pptx")
truth = {"overflowing": [{"slide": 1, "shape": "Box narrow"}, {"slide": 1, "shape": "Unwrapped line"},
                         {"slide": 2, "shape": "Title 1"}, {"slide": 2, "shape": "Content Placeholder 2"},
                         {"slide": 3, "shape": "Overflow table"}],
         "fitting": [{"slide": 1, "shape": "Title 1"}, {"slide": 1, "shape": "Box roomy"},
                     {"slide": 1, "shape": "Box near miss"}, {"slide": 1, "shape": "Badge"},
                     {"slide": 3, "shape": "Title 1"}, {"slide": 3, "shape": "Fit table"},
                     {"slide": 3, "shape": "Growing table"},
                     {"slide": 3, "shape": "Footnote"}, {"slide": 3, "shape": "Source"}]}
(out / "truth.json").write_text(json.dumps(truth, indent=1))
