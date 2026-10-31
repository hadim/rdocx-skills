"""D04 oracle: placeholders filled, the item row repeated 6 times with its formatting, total row last."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg, norm  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
items = [("Steel bolts M8", "120", "0.35"), ("Nuts & washers kit", "40", "4.10"), ("Hinge, stainless", "16", "7.80"),
         ("Wood screws 4x40", "500", "0.04"), ("Wall plugs 6 mm", "300", "0.03"), ("Cable ties 200 mm", "250", "0.05")]
all_text = "\n".join(pkg.text(p) for p in pkg.paragraphs())
ck.require("{{" not in all_text and "}}" not in all_text, "placeholders left in the document")
for label, value, flag in (("Client: ", "Acme Corp", None), ("Date: ", "31 October 2026", "i")):
    p = pkg.find_para(lambda t, label=label: t.startswith(label))
    if ck.require(p is not None and pkg.text(p) == label + value, f"no paragraph '{label}{value}'"):
        runs = pkg.runs_with_text(p)
        ck.require(all(pkg.run_flag(r, "b") for r, t in runs if t and label.startswith(t.strip() or "#")),
                   f"label {label!r} lost its bold")
        val_runs = [(r, t) for r, t in runs if t and t in value]
        ck.require(not any(pkg.run_flag(r, "b") for r, _ in val_runs), f"value of {label!r} became bold")
        if flag:
            ck.require(all(pkg.run_flag(r, flag) for r, _ in val_runs), f"value of {label!r} lost its italic")
ck.require(pkg.find_para(lambda t: t == "Thank you for your business, Acme Corp.") is not None,
           "closing paragraph not filled")
tbl = next(iter(pkg.body.iter(W + "tbl")), None)
if ck.require(tbl is not None, "no table"):
    rows = tbl.findall(W + "tr")
    got = [[norm(pkg.text(c)) for c in tr.findall(W + "tc")] for tr in rows]
    expected = [["Item", "Qty", "Price"]] + [list(i) for i in items] + [["Total", "", "471.30"]]
    ck.require(got == expected, f"table rows differ: {got}")
    if got == expected:
        for tr in rows[1:7]:
            cells = tr.findall(W + "tc")
            for c in cells[1:]:
                jc = c.find(f"{W}p/{W}pPr/{W}jc")
                ck.require(jc is not None and jc.get(W + "val") in ("right", "end"), "an item cell lost its right alignment")
            sizes = {r.find(f"{W}rPr/{W}sz").get(W + "val") if r.find(f"{W}rPr/{W}sz") is not None else None
                     for r in tr.iter(W + "r") if pkg.text(r)}
            ck.require(sizes == {"20"}, f"an item row's runs are not 10 pt: {sizes}")
        ck.require(all(pkg.run_flag(r, "b") for r, t in pkg.runs_with_text(rows[7].findall(W + "tc")[2]) if t),
                   "the total lost its bold")
ck.done()
