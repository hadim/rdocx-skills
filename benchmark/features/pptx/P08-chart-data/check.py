"""P08 oracle: the chart on slide 3 has 2 series whose names, category and value caches hold the new data (5
points); each formula (c:f) points at workbook cells that hold the same values; series fills, data labels (with
their number format), title and legend position are those of the input chart."""
import io
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402
from lxml import etree  # noqa: E402

S = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
     "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
     "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


def chart_of(pkg, slide):
    gf = oracle.find_one(pkg, slide, lambda sh: sh.find(".//c:chart", NS) is not None)
    return pkg.rel(slide, gf.find(".//c:chart", NS).get(oracle.q("r:id")))


def workbook(pkg, chart):
    """{sheet name: {cell ref: value}} of the chart's embedded workbook."""
    part = (pkg.rels_of_type(chart, "package") or [None])[0]
    z = zipfile.ZipFile(io.BytesIO(pkg.read(part)))
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(si.itertext()) for si in etree.fromstring(z.read("xl/sharedStrings.xml"))]
    out = {}
    for sh in wb.iter(f"{{{S['s']}}}sheet"):
        target = rels[sh.get(f"{{{S['r']}}}id")].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        cells = {}
        for c in etree.fromstring(z.read(target)).iter(f"{{{S['s']}}}c"):
            v, t = c.findtext("s:v", namespaces=S), c.get("t")
            if t == "s":
                v = shared[int(v)]
            elif t == "inlineStr":
                v = "".join(c.find("s:is", S).itertext())
            elif v is not None and t != "str":
                v = float(v)
            cells[c.get("r")] = v
        out[sh.get("name")] = cells
    return out


def cells(ref):
    sheet, rng = ref.replace("$", "").rsplit("!", 1)
    a, _, b = rng.partition(":")
    b = b or a
    (c1, r1), (c2, r2) = re.match(r"([A-Z]+)(\d+)", a).groups(), re.match(r"([A-Z]+)(\d+)", b).groups()
    col = lambda s: sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(s)))  # noqa: E731
    name = lambda n: (name((n - 1) // 26) if n > 26 else "") + chr(65 + (n - 1) % 26)  # noqa: E731
    return sheet.strip("'"), [f"{name(c)}{r}" for c in range(col(c1), col(c2) + 1) for r in range(int(r1), int(r2) + 1)]


def cache(el):
    return [pt.findtext("c:v", namespaces=NS) for pt in sorted(el.findall("c:pt", NS), key=lambda p: int(p.get("idx")))]


def look(pkg, chart):
    root = pkg.xml(chart)
    sers = root.findall(".//c:barChart/c:ser", NS)
    fill = [(s.find("c:spPr/a:solidFill/a:srgbClr", NS) is not None and s.find("c:spPr/a:solidFill/a:srgbClr", NS)
             .get("val")) for s in sers]
    dl = root.find(".//c:barChart/c:dLbls", NS)
    if dl is None and sers:
        dl = sers[0].find("c:dLbls", NS)
    labels = None if dl is None else (dl.find("c:showVal", NS).get("val"),
                                      (dl.find("c:numFmt", NS).get("formatCode") if dl.find("c:numFmt", NS) is not None
                                       else None))
    title = "".join(root.find("c:chart/c:title", NS).itertext()) if root.find("c:chart/c:title", NS) is not None else None
    legend = root.find("c:chart/c:legend/c:legendPos", NS)
    return {"fills": fill, "labels": labels, "title": title, "legend": None if legend is None else legend.get("val")}


def check(out, inputs):
    new = json.loads((inputs / "new_data.json").read_text())
    src = oracle.Pkg(inputs / "deck.pptx")
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    chart = chart_of(pkg, pkg.slides()[2])
    sers = pkg.xml(chart).findall(".//c:barChart/c:ser", NS)
    if len(sers) != 2:
        return errors + [f"{len(sers)} series"]
    try:
        wb = workbook(pkg, chart)
    except Exception as e:
        return errors + [f"embedded workbook unreadable: {e}"]
    for ser in sers:
        name = cache(ser.find("c:tx/c:strRef/c:strCache", NS))[0]
        cats = cache(ser.find("c:cat//c:strCache", NS)) if ser.find("c:cat//c:strCache", NS) is not None else None
        vals = [float(v) for v in cache(ser.find("c:val/c:numRef/c:numCache", NS))]
        if name not in new["series"]:
            errors.append(f"series name {name!r}")
            continue
        if cats != new["categories"]:
            errors.append(f"{name}: categories cache {cats}")
        if vals != [float(v) for v in new["series"][name]]:
            errors.append(f"{name}: values cache {vals}")
        for kind, path, want in (("name", "c:tx/c:strRef/c:f", [name]), ("categories", "c:cat//c:f", new["categories"]),
                                 ("values", "c:val/c:numRef/c:f", [float(v) for v in new["series"][name]])):
            f = ser.findtext(path, namespaces=NS)
            sheet, refs = cells(f)
            got = [wb.get(sheet, {}).get(r) for r in refs]
            if got != want:
                errors.append(f"{name}: {kind} formula {f} reads {got} in the workbook, expected {want}")
    before, after = look(src, chart_of(src, src.slides()[2])), look(pkg, chart)
    for k in before:
        if before[k] != after[k]:
            errors.append(f"formatting {k}: {after[k]!r}, was {before[k]!r}")
    return errors, {"workbook_sheets": list(wb)}


oracle.main(check)
