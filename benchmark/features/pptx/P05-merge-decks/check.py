"""P05 oracle: order A1, B2, B3, A2, A3, A4; the copies of B2 and B3 use layouts named as in B ("Harbor Title
Only", "Harbor Picture Caption") whose master's theme has Harbor's accent1 and major font, while A's slides keep
Northwind's; the B2 copy has a chart with B's categories and values and an embedded workbook; the B3 copy shows
B's picture. The presentation's masters are listed once each, master and layout ids are unique, there is at most
one notes master, and relationships, content types and reachability are clean."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS, q  # noqa: E402


def look(pkg, slide):
    layout = pkg.layout_of(slide)
    theme = pkg.theme_of(pkg.master_of(layout))
    t = pkg.xml(theme)
    acc = t.find(".//a:clrScheme/a:accent1/a:srgbClr", NS)
    major = t.find(".//a:fontScheme/a:majorFont/a:latin", NS)
    return pkg.layout_name(layout), (acc.get("val") if acc is not None else None), major.get("typeface")


def chart_data(pkg, slide):
    gf = oracle.find_one(pkg, slide, lambda sh: sh.find(".//c:chart", NS) is not None)
    if gf is None:
        return None, None
    chart = pkg.rel(slide, gf.find(".//c:chart", NS).get(q("r:id")))
    data = [([pt.findtext("c:v", namespaces=NS) for pt in s.iterfind("c:cat//c:pt", NS)],
             [pt.findtext("c:v", namespaces=NS) for pt in s.iterfind("c:val//c:pt", NS)])
            for s in pkg.xml(chart).iterfind(".//c:ser", NS)]
    return chart, data


def check(out, inputs):
    b = oracle.Pkg(inputs / "b.pptx")
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    slides = pkg.slides()
    order = [(pkg.title(s) or "?").split(":")[0] for s in slides]
    if order != ["A1", "B2", "B3", "A2", "A3", "A4"]:
        return errors + [f"order {order}"]
    harbor, northwind = ("0B6E4F", "Calibri"), ("1E2761", "Cambria")  # lib/inputs.py HARBOR and NORTHWIND
    for k, s in enumerate(slides):
        name, acc, major = look(pkg, s)
        want = harbor if order[k].startswith("B") else northwind
        if (acc, major) != want:
            errors.append(f"{order[k]}: theme accent1/major font {acc}/{major}, expected {want[0]}/{want[1]}")
        if order[k] == "B2" and name != "Harbor Title Only":
            errors.append(f"B2 copy on layout {name!r}")
        if order[k] == "B3" and name != "Harbor Picture Caption":
            errors.append(f"B3 copy on layout {name!r}")
    chart, data = chart_data(pkg, slides[1])
    _, bdata = chart_data(b, b.slides()[1])
    if data != bdata:
        errors.append(f"B2 copy chart data {data}, expected {bdata}")
    elif not pkg.rels_of_type(chart, "package"):
        errors.append("B2 copy chart has no embedded workbook")
    photo = (inputs / "harbor.png").read_bytes()
    if photo not in [pkg.read(t) for t in pkg.rels_of_type(slides[2], "image")]:
        errors.append("B3 copy does not show B's picture")
    pres = pkg.xml(pkg.main)
    masters = [(m.get("id"), pkg.rel(pkg.main, m.get(q("r:id")))) for m in pres.iterfind("p:sldMasterIdLst/p:sldMasterId", NS)]
    if len({m for _, m in masters}) != len(masters) or set(m for _, m in masters) != set(pkg.rels_of_type(pkg.main, "slideMaster")):
        errors.append("slide masters listed and related disagree")
    used = {pkg.master_of(pkg.layout_of(s)) for s in slides}
    if not used <= {m for _, m in masters}:
        errors.append("a slide uses a master that sldMasterIdLst does not list")
    all_ids = [int(i) for i, _ in masters]
    for _, m in masters:
        all_ids += [int(e.get("id")) for e in pkg.xml(m).iterfind("p:sldLayoutIdLst/p:sldLayoutId", NS)]
    if len(all_ids) != len(set(all_ids)) or min(all_ids) < 2147483648:
        errors.append(f"master and layout ids are not unique or below 2^31: {sorted(all_ids)}")
    if len(pres.findall("p:notesMasterIdLst/p:notesMasterId", NS)) > 1:
        errors.append("more than one notes master")
    orphans = pkg.orphans()
    if orphans:
        errors.append(f"orphan parts: {orphans}")
    return errors, {"masters": len(masters), "b2_notes_kept": "B2 notes" in (pkg.notes_text(slides[1]) or "")}


oracle.main(check)
