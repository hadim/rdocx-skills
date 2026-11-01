# bench: tools=python-pptx, lxml; processes=1; raw_xml=yes
import json
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
src, out = sys.argv[1], sys.argv[2]
content = json.load(open(src))
prs = Presentation()  # the default template is 4:3: set 16:9
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
theme_part = prs.slide_master.part.part_related_by(RT.THEME)  # python-pptx has no theme API: edit its XML
theme = etree.fromstring(theme_part.blob)
theme.set("name", content["theme"]["name"])
scheme = theme.find(f".//{{{A}}}clrScheme")
scheme.set("name", content["theme"]["name"])
for key, value in content["theme"]["colors"].items():
    slot = scheme.find(f"{{{A}}}{key}")
    slot.clear()
    etree.SubElement(slot, f"{{{A}}}srgbClr", val=value)
theme.find(f".//{{{A}}}majorFont/{{{A}}}latin").set("typeface", content["theme"]["headFont"])
theme.find(f".//{{{A}}}minorFont/{{{A}}}latin").set("typeface", content["theme"]["bodyFont"])
theme_part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)
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
        data = CategoryChartData()
        data.categories = s["chart"]["categories"]
        for ser in s["chart"]["series"]:
            data.add_series(ser["name"], ser["values"])
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.6), Inches(12.1),
                                       Inches(5.2), data).chart
        chart.has_legend, chart.legend.position = True, XL_LEGEND_POSITION.BOTTOM
    if s["kind"] == "picture":
        slide.placeholders[1].insert_picture(os.path.join(os.path.dirname(src), s["image"]))
        slide.placeholders[2].text = s["caption"]
    slide.notes_slide.notes_text_frame.text = s["notes"]
    if s["kind"] != "title":  # slide numbers: python-pptx has no API; clone the layout's placeholder, add the field
        lph = next(p for p in slide.slide_layout.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
        slide.shapes.clone_placeholder(lph)
        p = slide.placeholders[lph.placeholder_format.idx].text_frame.paragraphs[0]._p
        p.append(etree.fromstring(f'<a:fld xmlns:a="{A}" id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" '
                                  f'type="slidenum"><a:rPr lang="en-US"/><a:t>&#8249;#&#8250;</a:t></a:fld>'))
prs.save(out)
