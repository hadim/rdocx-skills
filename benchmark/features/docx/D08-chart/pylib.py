import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.chart.xmlwriter import ChartXmlWriter
from pptx.enum.chart import XL_CHART_TYPE

src, out = sys.argv[1:3]
doc = Document(src)
data = CategoryChartData()
data.categories = ["Q1", "Q2", "Q3", "Q4"]
data.add_series("2025", (10, 12, 14, 9))
data.add_series("2026", (11, 15, 13, 17))
xml = ChartXmlWriter(XL_CHART_TYPE.COLUMN_CLUSTERED, data).xml
workbook = Part(PackURI("/word/embeddings/Microsoft_Excel_Worksheet1.xlsx"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", data.xlsx_blob, doc.part.package)
chart_xml = parse_xml(xml.encode())
chart = Part(PackURI("/word/charts/chart1.xml"), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
             b"", doc.part.package)
wb_rid = chart.relate_to(workbook, RT.PACKAGE)
chart_xml.append(parse_xml(f'<c:externalData {nsdecls("c", "r")} r:id="{wb_rid}"><c:autoUpdate val="0"/></c:externalData>'))
chart._blob = etree.tostring(chart_xml, xml_declaration=True, encoding="UTF-8", standalone=True)
rid = doc.part.relate_to(chart, RT.CHART)
anchor = next(i for i, p in enumerate(doc.paragraphs) if p.text == "Chart anchor: sales by quarter.")
holder = doc.paragraphs[anchor + 1].insert_paragraph_before()
holder._p.append(parse_xml(
    f'<w:r {nsdecls("w", "wp", "a", "c", "r")}><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="5400000" cy="2880000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="100" name="Chart 1"/><wp:cNvGraphicFramePr/><a:graphic>'
    f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="{rid}"/>'
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'))
doc.save(out)
