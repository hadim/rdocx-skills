"""D08 oracle: a chart part (clustered columns, 2 series x 4 categories, right values) drawn in the paragraph
after the anchor, with an embedded workbook so that Word can edit its data."""
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import NS, W, Check, Pkg, q  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
body = [p for p in pkg.body if p.tag == W + "p"]
texts = [pkg.text(p) for p in body]
k = texts.index("Chart anchor: sales by quarter.") if "Chart anchor: sales by quarter." in texts else None
chart = None
if ck.require(k is not None and k + 1 < len(body), "anchor paragraph missing"):
    holders = [p for p in body[k:k + 2] if p.find(".//c:chart", NS) is not None]
    if ck.require(holders, "no chart in or right after the anchor paragraph"):
        rid = holders[0].find(".//c:chart", NS).get(q("r:id"))
        rel = pkg.rels(pkg.main).get(rid)
        ck.require(rel and rel[0] == "chart", f"the chart's r:id {rid} is not a chart relationship")
        chart = rel[1] if rel else None
if chart:
    ck.require((pkg.content_type(chart) or "").endswith("drawingml.chart+xml"), "wrong chart content type")
    root = pkg.xml(chart)
    bar = root.find(".//c:barChart", NS)
    if ck.require(bar is not None, "no c:barChart"):
        ck.require(bar.find("c:barDir", NS).get("val") == "col", "bars are not columns")
        ck.require(bar.find("c:grouping", NS).get("val") == "clustered", "columns are not clustered")
        series = bar.findall("c:ser", NS)

        def pts(el):
            return [p.findtext("c:v", namespaces=NS) for p in el.iter(q("c:pt"))] if el is not None else []

        got = [("".join(t.text or "" for t in s.find("c:tx", NS).iter(q("c:v"))), pts(s.find("c:cat", NS)),
                [float(v) for v in pts(s.find("c:val", NS))]) for s in series]
        expected = [("2025", ["Q1", "Q2", "Q3", "Q4"], [10, 12, 14, 9]), ("2026", ["Q1", "Q2", "Q3", "Q4"], [11, 15, 13, 17])]
        ck.require(got == expected, f"series differ: {got}")
    ext = root.find("c:externalData", NS)
    if ck.require(ext is not None, "no c:externalData: Word cannot edit the chart's data"):
        target = pkg.rels(chart).get(ext.get(q("r:id")))
        ok = target and target[1] in pkg.names and zipfile.is_zipfile(io.BytesIO(pkg.read(target[1])))
        ck.require(ok, "the chart's embedded workbook is missing")
        if ok:
            xlsx = zipfile.ZipFile(io.BytesIO(pkg.read(target[1])))
            sheet = next((n for n in xlsx.namelist() if n.startswith("xl/worksheets/sheet")), None)
            ck.require(sheet is not None, "the embedded workbook has no sheet")
ck.done()
