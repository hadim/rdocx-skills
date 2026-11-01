import copy
import sys

from lxml import etree

path = sys.argv[1]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
tree = etree.parse(path)
tbl = tree.find(f".//{W}tbl")


def text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def set_width(tc, twips):
    tc.find(f"{W}tcPr/{W}tcW").set(W + "w", str(twips))


# 1. a Supplier column after the first one: copy the first cell of each row, keep its first run's format
grid = tbl.find(W + "tblGrid")
col = copy.deepcopy(grid[0])
col.set(W + "w", "1701")
grid[0].addnext(col)
for tr, value in zip(tbl.findall(W + "tr"), ["Supplier", "Acme", "Bolt & Co", "Acme", ""]):
    first = tr.findall(W + "tc")[0]
    new = copy.deepcopy(first)
    set_width(new, 1701)
    runs = new.findall(f"{W}p/{W}r")
    for r in runs[1:] if value else runs:
        r.getparent().remove(r)
    if value:
        runs[0].find(W + "t").text = value
    first.addnext(new)
# 2. merge 'Total' with the cell to its right
total_row = next(tr for tr in tbl.findall(W + "tr") if text(tr.findall(W + "tc")[0]) == "Total")
total, right = total_row.findall(W + "tc")[:2]
span = etree.SubElement(total.find(W + "tcPr"), W + "gridSpan")
span.set(W + "val", "2")
set_width(total, int(total.find(f"{W}tcPr/{W}tcW").get(W + "w")) + 1701)
total_row.remove(right)
# 3. delete the 'Nuts' row
tbl.remove(next(tr for tr in tbl.findall(W + "tr") if text(tr.findall(W + "tc")[0]) == "Nuts"))
# 4. shade the 'Amount' header cell (w:shd follows w:tcW / w:gridSpan in w:tcPr)
amount = next(tc for tc in tbl.findall(W + "tr")[0].findall(W + "tc") if text(tc) == "Amount")
shd = etree.Element(W + "shd")
shd.set(W + "val", "clear")
shd.set(W + "color", "auto")
shd.set(W + "fill", "FFC000")
amount.find(f"{W}tcPr/{W}tcW").addnext(shd)
tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
