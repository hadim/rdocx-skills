"""P10 oracle: the slide-2 table has 5 rows of 4 cells on a 4-column grid; the new last row reads Total / (empty) /
2,980 / +5 %, its first cell spans 2 columns (gridSpan=2, the next cell hMerge=1), it has the height of the row
above and the same run size; cell (2,3) has a solid FFC000 fill; every other cell keeps its text and fill."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402
from lxml import etree  # noqa: E402


def table(pkg):
    slide = pkg.slides()[1]
    return pkg.xml(slide).find(".//a:tbl", NS)


def cell_text(tc):
    return "\n".join(oracle.Pkg.para_text(p) for p in tc.iter(oracle.q("a:p")))


def fill(tc):
    tcpr = tc.find("a:tcPr", NS)
    el = None if tcpr is None else next((c for c in tcpr if etree.QName(c).localname.endswith("Fill")), None)
    return oracle.canon(el)


def sizes(tc):
    return {r.get("sz") for r in tc.iter(oracle.q("a:rPr")) if r.getparent().tag == oracle.q("a:r")}


def check(out, inputs):
    before, pkg = table(oracle.Pkg(inputs / "deck.pptx")), oracle.Pkg(out)
    errors = pkg.integrity()
    tbl = table(pkg)
    rows = tbl.findall("a:tr", NS)
    if len(tbl.findall("a:tblGrid/a:gridCol", NS)) != 4:
        errors.append("grid is no longer 4 columns")
    if len(rows) != 5 or any(len(r.findall("a:tc", NS)) != 4 for r in rows):
        return errors + [f"{len(rows)} rows, cells per row {[len(r.findall('a:tc', NS)) for r in rows]}"]
    last, above = rows[4].findall("a:tc", NS), rows[3].findall("a:tc", NS)
    texts = [cell_text(tc).strip() for tc in last]
    if texts[0] != "Total" or texts[2:] != ["2,980", "+5 %"] or texts[1]:
        errors.append(f"new row reads {texts}")
    if last[0].get("gridSpan") != "2" or last[1].get("hMerge") != "1":
        errors.append(f"cells not merged (gridSpan {last[0].get('gridSpan')}, hMerge {last[1].get('hMerge')})")
    if rows[4].get("h") != rows[3].get("h"):
        errors.append(f"new row height {rows[4].get('h')}, row above {rows[3].get('h')}")
    for k in (0, 2, 3):
        if sizes(last[k]) != sizes(above[k]):
            errors.append(f"new row cell {k} run size {sizes(last[k])}, row above {sizes(above[k])}")
    target = rows[2].findall("a:tc", NS)[3]
    sf = target.find("a:tcPr/a:solidFill/a:srgbClr", NS)
    if sf is None or sf.get("val").upper() != "FFC000":
        errors.append("cell (2,3) has no solid FFC000 fill")
    for r, (br, ar) in enumerate(zip(before.findall("a:tr", NS), rows)):
        for c, (bc, ac) in enumerate(zip(br.findall("a:tc", NS), ar.findall("a:tc", NS))):
            if cell_text(bc) != cell_text(ac):
                errors.append(f"cell ({r},{c}) text changed")
            if (r, c) != (2, 3) and fill(bc) != fill(ac):
                errors.append(f"cell ({r},{c}) fill changed")
    return errors


oracle.main(check)
