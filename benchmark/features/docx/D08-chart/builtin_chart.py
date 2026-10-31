import os
import re
import sys

from openpyxl import Workbook

root = sys.argv[1]
CATS = ["Q1", "Q2", "Q3", "Q4"]
SERIES = [("2025", [10, 12, 14, 9]), ("2026", [11, 15, 13, 17])]
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def path(*parts):
    return os.path.join(root, *parts)


def edit(name, old, new):
    with open(path(name), encoding="utf-8") as f:
        xml = f.read()
    with open(path(name), "w", encoding="utf-8") as f:
        f.write(xml.replace(old, new, 1))


# the embedded workbook Word edits the data in
os.makedirs(path("word", "embeddings"), exist_ok=True)
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.append([""] + [name for name, _ in SERIES])
for i, cat in enumerate(CATS):
    ws.append([cat] + [values[i] for _, values in SERIES])
wb.save(path("word", "embeddings", "Microsoft_Excel_Worksheet1.xlsx"))

# the chart part
cols = "BCDEFG"
pts = lambda vals: "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
series = ""
for k, (name, values) in enumerate(SERIES):
    series += (
        f'<c:ser><c:idx val="{k}"/><c:order val="{k}"/><c:tx><c:strRef><c:f>Sheet1!${cols[k]}$1</c:f><c:strCache>'
        f'<c:ptCount val="1"/>{pts([name])}</c:strCache></c:strRef></c:tx><c:invertIfNegative val="0"/>'
        f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${len(CATS) + 1}</c:f><c:strCache><c:ptCount val="{len(CATS)}"/>'
        f'{pts(CATS)}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!${cols[k]}$2:${cols[k]}$'
        f'{len(CATS) + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>'
        f'{pts(values)}</c:numCache></c:numRef></c:val></c:ser>')
chart = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace '
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R}"><c:roundedCorners val="0"/>'
    '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>'
    f'<c:grouping val="clustered"/><c:varyColors val="0"/>{series}<c:gapWidth val="150"/>'
    '<c:axId val="500000001"/><c:axId val="500000002"/></c:barChart><c:catAx><c:axId val="500000001"/>'
    '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/>'
    '<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
    '<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
    '<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
    '<c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/>'
    '<c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
    '</c:legend><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/>'
    '</c:externalData></c:chartSpace>')
os.makedirs(path("word", "charts", "_rels"), exist_ok=True)
with open(path("word", "charts", "chart1.xml"), "w", encoding="utf-8") as f:
    f.write(chart)
with open(path("word", "charts", "_rels", "chart1.xml.rels"), "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.'
            'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.'
            'openxmlformats.org/officeDocument/2006/relationships/package" '
            'Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/></Relationships>')
edit("[Content_Types].xml", "</Types>",
     '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
     '<Override PartName="/word/charts/chart1.xml" '
     'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')
edit("word/_rels/document.xml.rels", "</Relationships>",
     f'<Relationship Id="rIdChart1" Type="{R}/chart" Target="charts/chart1.xml"/></Relationships>')

# the paragraph that shows it, after the anchor
with open(path("word", "document.xml"), encoding="utf-8") as f:
    doc = f.read()
ids = [int(i) for i in re.findall(r'<wp:docPr id="(\d+)"', doc)]
pid = max(ids, default=0) + 1
para = (f'<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5400000" '
        f'cy="2880000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{pid}" name="Chart {pid}"/>'
        '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart '
        f'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/></a:graphicData>'
        '</a:graphic></wp:inline></w:drawing></w:r></w:p>')
anchor = re.search(r"<w:p[ >](?:(?!<w:p[ >]).)*?Chart anchor: .*?</w:p>", doc, re.S)
with open(path("word", "document.xml"), "w", encoding="utf-8") as f:
    f.write(doc[:anchor.end()] + para + doc[anchor.end():])
