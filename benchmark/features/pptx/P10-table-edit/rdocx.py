# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
shape = next(sh for sh in prs.slides[1].shapes if sh.has_table)
shape.table.rows.add_row()  # at the end, formatted like its neighbour; re-fetch the cells after it
table = shape.table
for c, text in enumerate(["Total", "", "2,980", "+5 %"]):
    table.cell(4, c).text = text  # the new cells carry the neighbour's run size
table.cell(4, 0).merge(table.cell(4, 1))
table.cell(2, 3).fill.solid()
table.cell(2, 3).fill.fore_color.rgb = "FFC000"
pptx_ops.save_atomic(prs, sys.argv[2], sys.argv[1])
