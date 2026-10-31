"""Create the corpus deck from its plan with rpptx (the rdocx stack's creation path).

    $R/python pptx_create_rdocx.py plan.json out.pptx      (pptx_ops on PYTHONPATH)

Same content as corpus/gen_pptx.js (pptxgenjs) on rpptx's default template (16:9, its eleven layouts): the
"Harbourside" theme colours and Cambria / Calibri fonts, one section per chapter, speaker notes on every slide,
title, section, bullet, two-column, picture, native chart (bar, line, pie), table, stat-callout and long-text
slides, footer text and slide numbers on every slide but the title slide (PowerPoint's Header and Footer dialog).
rpptx writes no chart title and no data labels (no API): the chart's title goes into the slide's commentary.
"""
import json
import sys

import pptx_ops
import rpptx
from rpptx.enum.dml import MSO_THEME_COLOR as TC
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.enum.text import PP_ALIGN
from rpptx.util import Inches, Pt

plan_path, out = sys.argv[1:3]
plan = json.load(open(plan_path, encoding="utf-8"))
THEME = {"dk1": "1B2A3A", "lt1": "FFFFFF", "dk2": "0B4F6C", "lt2": "EEF3F6", "accent1": "0B7A75",
         "accent2": "E07A1F", "accent3": "5B8E7D", "accent4": "8C5383", "accent5": "C9A227", "accent6": "3D6FB6",
         "hlink": "0B4F6C", "folHlink": "8C5383"}

prs = rpptx.Presentation()  # 16:9, 13.33 x 7.5 in
theme = prs.slide_master.theme
theme.name = "Harbourside"
for key, value in THEME.items():
    theme.colors[key] = value
theme.fonts.major.latin, theme.fonts.minor.latin = "Cambria", "Calibri"
prs.core_properties.title = plan["meta"]["title"]
prs.core_properties.author = "Corporate Affairs"
prs.app_properties.company = plan["meta"]["company"]
layout = {lay.name: lay for lay in prs.slide_layouts}


def fill(text_frame, items, size=None, head=None, head_color=None):
    """One paragraph per item; an optional bold, coloured head line first."""
    lines = ([head] if head else []) + list(items)
    text_frame.text = lines[0]
    for item in lines[1:]:
        text_frame.add_paragraph().text = item
    for k, p in enumerate(text_frame.paragraphs):
        for r in p.runs:
            if size:
                r.font.size = Pt(size)
            if head and k == 0:
                r.font.bold = True
                r.font.color.theme_color = head_color


sections, kinds = [], {"bar": rpptx.XL_CHART_TYPE.COLUMN_CLUSTERED, "line": rpptx.XL_CHART_TYPE.LINE,
                       "pie": rpptx.XL_CHART_TYPE.PIE}
for k, s in enumerate(plan["slides"]):
    t = s["type"]
    if not sections or sections[-1][0] != s["section"]:
        sections.append((s["section"], []))
    sections[-1][1].append(k)
    if t == "title":
        slide = prs.slides.add_slide(layout["Title Slide"])
        slide.shapes.title.text = s["title"]
        slide.placeholders[1].text = s["subtitle"]
    elif t == "section":
        slide = prs.slides.add_slide(layout["Section Header"])
        slide.shapes.title.text = s["title"]
        slide.placeholders[1].text = s["subtitle"]
    elif t in ("bullets", "closing"):
        slide = prs.slides.add_slide(layout["Title and Content"])
        slide.shapes.title.text = s["title"]
        fill(slide.placeholders[1].text_frame, s["bullets"])
    elif t == "long_text":
        slide = prs.slides.add_slide(layout["Title and Content"])
        slide.shapes.title.text = s["title"]
        fill(slide.placeholders[1].text_frame, s["text"], size=14)
    elif t == "bullets_pic":
        slide = prs.slides.add_slide(layout["Two Content"])
        slide.shapes.title.text = s["title"]
        fill(slide.placeholders[1].text_frame, s["bullets"], size=16)
        slide.shapes.remove(slide.placeholders[2])
        w = 5.6
        h = min(5.0, w * s["image"]["h"] / s["image"]["w"])
        pic = slide.shapes.add_picture(s["image"]["path"], Inches(7.2), Inches(1.6), Inches(w), Inches(h))
        pic.name, pic.alt_text = "Figure", "Chart from the report"
    elif t == "two_col":
        slide = prs.slides.add_slide(layout["Two Content"])
        slide.shapes.title.text = s["title"]
        fill(slide.placeholders[1].text_frame, s["bullets"], size=15, head=s["left_head"], head_color=TC.ACCENT_1)
        fill(slide.placeholders[2].text_frame, s["right"], size=15, head=s["right_head"], head_color=TC.ACCENT_2)
    elif t == "chart":
        slide = prs.slides.add_slide(layout["Title Only"])
        slide.shapes.title.text = s["title"]
        spec = s["chart"]
        data = rpptx.CategoryChartData()
        data.categories = spec["series"][0]["labels"]
        for ser in spec["series"]:
            data.add_series(ser["name"], ser["values"])
        frame = slide.shapes.add_chart(kinds[spec["type"]], Inches(0.6), Inches(1.45), Inches(8.0), Inches(5.3), data)
        if len(spec["series"]) > 1 or spec["type"] == "pie":
            frame.chart.has_legend = True
            frame.chart.legend_position = "b"
        box = slide.shapes.add_textbox(Inches(8.9), Inches(1.6), Inches(3.8), Inches(5.1))
        box.text_frame.word_wrap = True
        fill(box.text_frame, s["commentary"], size=15, head=spec["title"], head_color=TC.TEXT_2)
    elif t == "table":
        slide = prs.slides.add_slide(layout["Title Only"])
        slide.shapes.title.text = s["title"]
        rows = s["table"]
        frame = slide.shapes.add_table(len(rows), len(rows[0]), Inches(0.6), Inches(1.5), Inches(12.1),
                                       Inches(0.45 * len(rows)))
        frame.name = "Data table"
        table = frame.table
        for i, row in enumerate(rows):
            for j, text in enumerate(row):
                cell = table.cell(i, j)
                cell.text = text
                cell.fill.solid()
                cell.fill.fore_color.theme_color = TC.ACCENT_1 if i == 0 else TC.BACKGROUND_2 if i % 2 == 0 else TC.BACKGROUND_1
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if j == 0 else (PP_ALIGN.CENTER if i == 0 else PP_ALIGN.RIGHT)
                for r in p.runs:
                    r.font.size = Pt(14)
                    if i == 0:
                        r.font.bold = True
                        r.font.color.theme_color = TC.BACKGROUND_1
    elif t == "stats":
        slide = prs.slides.add_slide(layout["Title Only"])
        slide.shapes.title.text = s["title"]
        for n, st in enumerate(s["stats"]):
            x = Inches(0.6 + n * 4.1)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.7), Inches(3.8), Inches(3.2))
            card.name = f"Stat card {n + 1}"
            card.fill.solid()
            card.fill.fore_color.theme_color = TC.BACKGROUND_2
            card.line.fill.background()
            value = slide.shapes.add_textbox(x, Inches(2.0), Inches(3.8), Inches(1.4))
            value.name = f"Stat value {n + 1}"
            value.text_frame.text = st["value"]
            p = value.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.runs[0].font.size, p.runs[0].font.bold = Pt(60), True
            p.runs[0].font.color.theme_color = TC.ACCENT_1
            label = slide.shapes.add_textbox(x + Inches(0.2), Inches(3.5), Inches(3.4), Inches(1.2))
            label.name = f"Stat label {n + 1}"
            label.text_frame.word_wrap = True
            label.text_frame.text = st["label"]
            label.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            label.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
        cap = slide.shapes.add_textbox(Inches(0.6), Inches(5.3), Inches(12.1), Inches(1.2))
        cap.name = "Caption"
        cap.text_frame.word_wrap = True
        cap.text_frame.text = s["caption"]
        run = cap.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.italic = Pt(16), True
        run.font.color.theme_color = TC.TEXT_2
    else:
        raise ValueError(f"unknown slide type {t}")
    slide.notes_text = s["notes"]

prs.set_sections(sections)
prs.set_header_footer(slide_number=True, footer=plan["meta"]["footer"])  # all but title slides, as the dialog
pptx_ops.save_atomic(prs, out)
