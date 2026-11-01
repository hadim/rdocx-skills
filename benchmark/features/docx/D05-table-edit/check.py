"""D05 oracle: the table after the four edits (column, horizontal merge, row deletion, shading)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg, norm  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
tables = list(pkg.body.iter(W + "tbl"))
if ck.require(len(tables) == 1, f"{len(tables)} tables, expected 1"):
    tbl = tables[0]
    grid = tbl.findall(f"{W}tblGrid/{W}gridCol")
    ck.require(len(grid) == 5, f"table grid has {len(grid)} columns, expected 5")
    rows = tbl.findall(W + "tr")
    got = [[norm(pkg.text(c)) for c in tr.findall(W + "tc")] for tr in rows]
    expected = [["Item", "Supplier", "Qty", "Price", "Amount"], ["Bolts", "Acme", "10", "0.50", "5.00"],
                ["Washers", "Acme", "30", "0.10", "3.00"], ["Total", "", "", "12.00"]]
    ck.require(got == expected, f"table cells differ: {got}")

    def span(tc):
        g = tc.find(f"{W}tcPr/{W}gridSpan")
        return int(g.get(W + "val")) if g is not None else 1

    for i, tr in enumerate(rows):
        ck.require(sum(span(tc) for tc in tr.findall(W + "tc")) == 5, f"row {i} does not span 5 grid columns")
    if got == expected:
        total = rows[3].findall(W + "tc")[0]
        ck.require(span(total) == 2 and total.find(f"{W}tcPr/{W}vMerge") is None, "'Total' does not span 2 columns")
        # Word drops the paragraph of an empty cell it merges in; an extra empty paragraph doubles the row's height
        ck.require(len(total.findall(W + "p")) == 1, f"the merged cell holds {len(total.findall(W + 'p'))} paragraphs")
        head = rows[0].findall(W + "tc")
        for c in (0, 1, 4):
            ck.require(any(pkg.run_flag(r, "b") for r, t in pkg.runs_with_text(head[c]) if t),
                       f"header {got[0][c]!r} is not bold")
        fills = [(tc.find(f"{W}tcPr/{W}shd").get(W + "fill") or "").upper() if tc.find(f"{W}tcPr/{W}shd") is not None
                 else None for tc in head]
        ck.require(fills[4] == "FFC000", f"'Amount' header fill is {fills[4]}")
        ck.require(fills.count("FFC000") == 1, "other header cells were shaded")
    ck.info(f"grid width {sum(int(g.get(W + 'w') or 0) for g in grid)} twips (text width 9071)")
ck.require(not any("Nuts" in pkg.text(p) for p in pkg.paragraphs()), "'Nuts' is still in the document")
ck.require(pkg.find_para(lambda t: t == "Prices are in euros.") is not None, "paragraph after the table lost")
ck.done()
