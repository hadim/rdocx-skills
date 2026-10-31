# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import json
import os
import sys

import pptx_ops
import rpptx
from rpptx.util import Inches

src, out = sys.argv[1], sys.argv[2]
content = json.load(open(src))
prs = rpptx.Presentation()  # 16:9, the eleven default layouts
theme = prs.slide_master.theme
theme.name = content["theme"]["name"]
for key, value in content["theme"]["colors"].items():
    theme.colors[key] = value
theme.fonts.major.latin, theme.fonts.minor.latin = content["theme"]["headFont"], content["theme"]["bodyFont"]
layout = {lay.name: lay for lay in prs.slide_layouts}
LAYOUT = {"title": "Title Slide", "section": "Section Header", "closing": "Section Header",
          "bullets": "Title and Content", "table": "Title Only", "chart": "Title Only", "two_columns": "Two Content",
          "picture": "Picture with Caption"}


def fill(text_frame, items):
    text_frame.text = items[0]
    for item in items[1:]:
        text_frame.add_paragraph().text = item


for s in content["slides"]:
    slide = prs.slides.add_slide(layout[LAYOUT[s["kind"]]])
    slide.shapes.title.text = s["title"]
    if "subtitle" in s:
        slide.placeholders[1].text = s["subtitle"]
    if s["kind"] == "bullets":
        fill(slide.placeholders[1].text_frame, s["bullets"])
    if s["kind"] == "two_columns":
        fill(slide.placeholders[1].text_frame, s["left"])
        fill(slide.placeholders[2].text_frame, s["right"])
    if s["kind"] == "table":
        rows = s["table"]
        table = slide.shapes.add_table(len(rows), len(rows[0]), Inches(0.6), Inches(1.7), Inches(12.1),
                                       Inches(0.5 * len(rows))).table
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                table.cell(r, c).text = text
    if s["kind"] == "chart":
        data = rpptx.CategoryChartData()
        data.categories = s["chart"]["categories"]
        for ser in s["chart"]["series"]:
            data.add_series(ser["name"], ser["values"])
        frame = slide.shapes.add_chart(rpptx.XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.6), Inches(12.1),
                                       Inches(5.2), data)
        frame.chart.legend_position = "b"
    if s["kind"] == "picture":
        slide.placeholders[1].insert_picture(os.path.join(os.path.dirname(src), s["image"]))
        slide.placeholders[2].text = s["caption"]
    slide.notes_text = s["notes"]
prs.set_header_footer(slide_number=True)  # every slide but title slides, as the Header and Footer dialog
pptx_ops.save_atomic(prs, out)
