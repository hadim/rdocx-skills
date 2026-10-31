import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
data = rdocx.CategoryChartData()
data.categories = ["Q1", "Q2", "Q3", "Q4"]
data.add_series("2025", (10, 12, 14, 9))
data.add_series("2026", (11, 15, 13, 17))
anchor = next(p for p in doc.paragraphs if p.text == "Chart anchor: sales by quarter.")
chart = doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, rdocx.Cm(15), rdocx.Cm(8), after=anchor)
doc.set_chart_legend(chart, visible=True, position="b")
docx_ops.save_atomic(doc, out, src)
