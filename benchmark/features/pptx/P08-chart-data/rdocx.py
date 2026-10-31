# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
chart = next(sh for sh in prs.slides[2].shapes if sh.has_chart).chart
data = rpptx.CategoryChartData()
data.categories = ["Q1 26", "Q2 26", "Q3 26", "Q4 26", "Q1 27"]
data.add_series("North", (1250, 1330, 1410, 1535, 1600))
data.add_series("South", (990, 1005, 1000, 1070, 1110))
chart.replace_data(data)
pptx_ops.save_atomic(prs, sys.argv[2], sys.argv[1])
