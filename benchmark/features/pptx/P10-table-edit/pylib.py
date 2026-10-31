# bench: tools=python-pptx, lxml; processes=1; raw_xml=yes
import copy
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor

prs = Presentation(sys.argv[1])
table = next(sh for sh in prs.slides[1].shapes if sh.has_table).table
tbl = table._tbl  # python-pptx cannot add a row: copy the last <a:tr>
tbl.append(copy.deepcopy(tbl.tr_lst[-1]))
for c, text in enumerate(["Total", "", "2,980", "+5 %"]):
    table.cell(4, c).text_frame.paragraphs[0].runs[0].text = text
table.cell(4, 0).merge(table.cell(4, 1))
table.cell(2, 3).fill.solid()
table.cell(2, 3).fill.fore_color.rgb = RGBColor(0xFF, 0xC0, 0x00)
prs.save(sys.argv[2])
