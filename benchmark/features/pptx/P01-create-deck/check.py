"""P01 oracle: 10 slides at 16:9; the slides' master uses a theme named Northwind (theme or colour-scheme name)
with the 12 colours and the two fonts of content.json; every title sits in a title placeholder and the deck uses
at least 5 layouts; bullets are one paragraph per item with a bullet (explicit or inherited from a body
placeholder); the two columns are two shapes, left of each other; the table, the native clustered-column chart
(caches and an embedded workbook) and the picture (photo.png's bytes) carry the content; each slide's notes read
as given; every slide but the first shows a slide number (a slidenum field on the slide, or on a non-placeholder
shape of its layout or master, as PowerPoint draws them) and the first shows none."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS, q  # noqa: E402


def shows_number(pkg, slide):
    if pkg.xml(slide).find(".//p:cSld//a:fld[@type='slidenum']", NS) is not None:
        return True
    layout = pkg.layout_of(slide)
    master = pkg.master_of(layout)
    for part in (layout, master):
        if part == master and (pkg.xml(layout).get("showMasterSp") == "0" or pkg.xml(slide).get("showMasterSp") == "0"):
            continue
        for sh in pkg.shapes(part):
            if pkg.ph(sh) is None and sh.find(".//a:fld[@type='slidenum']", NS) is not None:
                return True
    return False


def bulleted(pkg, slide, shape, p):
    ppr = p.find("a:pPr", NS)
    if ppr is not None and (ppr.find("a:buChar", NS) is not None or ppr.find("a:buAutoNum", NS) is not None):
        return True
    if ppr is not None and ppr.find("a:buNone", NS) is not None:
        return False
    ph = pkg.ph(shape)
    if ph is None or ph["type"] not in ("body", "obj"):
        return False
    layout = pkg.layout_of(slide)
    for sh in pkg.shapes(layout):
        lph = pkg.ph(sh)
        if lph and lph["idx"] == ph["idx"] and sh.find("p:txBody/a:lstStyle/a:lvl1pPr/a:buNone", NS) is not None:
            return False
    style = pkg.xml(pkg.master_of(layout)).find("p:txStyles/p:bodyStyle/a:lvl1pPr", NS)
    return style is not None and style.find("a:buChar", NS) is not None


def items_shape(pkg, slide, items):
    """The shape whose paragraphs are exactly `items` (empty paragraphs aside), or None."""
    for sh in pkg.shapes(slide):
        paras = [p for p in sh.iter(q("a:p")) if pkg.para_text(p).strip()]
        if [pkg.para_text(p).strip() for p in paras] == items:
            return sh, paras
    return None, None


def x_of(pkg, slide, sh):
    off = sh.find("p:spPr/a:xfrm/a:off", NS)
    if off is None and pkg.ph(sh):
        for lsh in pkg.shapes(pkg.layout_of(slide)):
            if pkg.ph(lsh) and pkg.ph(lsh)["idx"] == pkg.ph(sh)["idx"]:
                off = lsh.find("p:spPr/a:xfrm/a:off", NS)
    return int(off.get("x")) if off is not None else None


def check(out, inputs):
    content = json.loads((inputs / "content.json").read_text())
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    slides = pkg.slides()
    if len(slides) != 10:
        return errors + [f"{len(slides)} slides"]
    sz = pkg.xml(pkg.main).find("p:sldSz", NS)
    if abs(int(sz.get("cx")) / int(sz.get("cy")) - 16 / 9) > 0.01:
        errors.append(f"slide size {sz.get('cx')}x{sz.get('cy')} is not 16:9")
    theme = pkg.xml(pkg.theme_of(pkg.master_of(pkg.layout_of(slides[0]))))
    T = content["theme"]
    names = {theme.get("name"), theme.find(".//a:clrScheme", NS).get("name")}
    if T["name"] not in names:
        errors.append(f"theme named {names}")
    for k, v in T["colors"].items():
        el = theme.find(f".//a:clrScheme/a:{k}/a:srgbClr", NS)
        if el is None or el.get("val").upper() != v:
            errors.append(f"theme colour {k} is not {v}")
    fonts = (theme.find(".//a:majorFont/a:latin", NS).get("typeface"),
             theme.find(".//a:minorFont/a:latin", NS).get("typeface"))
    if fonts != (T["headFont"], T["bodyFont"]):
        errors.append(f"theme fonts {fonts}")
    if len({pkg.layout_of(s) for s in slides}) < 5:
        errors.append(f"only {len({pkg.layout_of(s) for s in slides})} layouts used")
    for n, (slide, s) in enumerate(zip(slides, content["slides"]), 1):
        if pkg.title(slide) != s["title"]:
            errors.append(f"slide {n}: title placeholder reads {pkg.title(slide)!r}")
        text = pkg.text(slide)
        for key in ("subtitle", "caption"):
            if key in s and s[key] not in text:
                errors.append(f"slide {n}: {key} missing")
        for key in ("bullets", "left", "right"):
            if key in s:
                sh, paras = items_shape(pkg, slide, s[key])
                if sh is None:
                    errors.append(f"slide {n}: {key} are not one paragraph per item in one shape")
                elif not all(bulleted(pkg, slide, sh, p) for p in paras):
                    errors.append(f"slide {n}: {key} without bullets")
        if s["kind"] == "two_columns":
            left, _ = items_shape(pkg, slide, s["left"])
            right, _ = items_shape(pkg, slide, s["right"])
            if left is not None and right is not None and not x_of(pkg, slide, left) < x_of(pkg, slide, right):
                errors.append(f"slide {n}: left column not left of the right one")
        if s["kind"] == "table":
            tbl = pkg.xml(slide).find(".//a:tbl", NS)
            rows = [] if tbl is None else [[pkg.text(slide, tc).strip() for tc in tr.findall("a:tc", NS)]
                                           for tr in tbl.findall("a:tr", NS)]
            if rows != s["table"]:
                errors.append(f"slide {n}: table {rows}")
        if s["kind"] == "chart":
            gf = oracle.find_one(pkg, slide, lambda sh: sh.find(".//c:chart", NS) is not None)
            if gf is None:
                errors.append(f"slide {n}: no native chart")
            else:
                chart = pkg.rel(slide, gf.find(".//c:chart", NS).get(q("r:id")))
                root = pkg.xml(chart)
                bar = root.find(".//c:barChart", NS)
                if bar is None or bar.find("c:barDir", NS).get("val") != "col" or \
                        bar.find("c:grouping", NS).get("val") != "clustered":
                    errors.append(f"slide {n}: not a clustered column chart")
                got = [{"name": ser.findtext("c:tx//c:v", namespaces=NS),
                        "cats": [pt.findtext("c:v", namespaces=NS) for pt in ser.iterfind("c:cat//c:pt", NS)],
                        "values": [float(pt.findtext("c:v", namespaces=NS)) for pt in ser.iterfind("c:val//c:pt", NS)]}
                       for ser in root.iterfind(".//c:ser", NS)]
                want = [{"name": x["name"], "cats": s["chart"]["categories"], "values": [float(v) for v in x["values"]]}
                        for x in s["chart"]["series"]]
                if got != want:
                    errors.append(f"slide {n}: chart data {got}")
                if not pkg.rels_of_type(chart, "package"):
                    errors.append(f"slide {n}: chart without an embedded workbook")
        if s["kind"] == "picture":
            photo = (inputs / s["image"]).read_bytes()
            if photo not in [pkg.read(t) for t in pkg.rels_of_type(slide, "image")]:
                errors.append(f"slide {n}: photo.png not shown")
        if (pkg.notes_text(slide) or "").strip() != s["notes"]:
            errors.append(f"slide {n}: notes {pkg.notes_text(slide)!r}")
        if shows_number(pkg, slide) != (n > 1):
            errors.append(f"slide {n}: slide number {'shown' if n == 1 else 'missing'}")
    return errors


oracle.main(check)
