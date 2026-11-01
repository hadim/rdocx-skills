"""D01 oracle: the new report has every element the task lists (zipfile + lxml only)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import NS, W, Check, Pkg, norm, picture_targets, png_size, q  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
sects = pkg.sect_prs()
pg = sects[-1].find(W + "pgSz") if sects else None
w, h = (int(pg.get(W + "w")), int(pg.get(W + "h"))) if pg is not None else (0, 0)
ck.require(11895 <= w <= 11915 and 16830 <= h <= 16845, f"page is not A4 portrait: {w} x {h} twips")

paras = pkg.paragraphs()
by_text = {}
for p in paras:
    by_text.setdefault(norm(pkg.text(p)), p)
title = by_text.get("Quarterly Report")
ck.require(title is not None and pkg.para_style_name(title) == "title", "no 'Quarterly Report' paragraph in the Title style")
for text, level in (("Introduction", 1), ("Scope", 2), ("Details", 3), ("Results", 1)):
    hs = [p for p in paras if norm(pkg.text(p)) == text and pkg.para_style_name(p) == f"heading {level}"]
    ck.require(hs, f"no heading {level} '{text}'")

intro = pkg.find_para(lambda t: t.startswith("This report has"))
if ck.require(intro is not None, "no 'This report has ...' paragraph"):
    ck.require(norm(pkg.text(intro)) == "This report has bold findings and italic remarks.", "intro text differs")
    runs = pkg.runs_with_text(intro)
    ck.require(any("bold findings" in t and pkg.run_flag(r, "b") for r, t in runs), "'bold findings' is not bold")
    ck.require(any("italic remarks" in t and pkg.run_flag(r, "i") for r, t in runs), "'italic remarks' is not italic")
    ck.require(not any("This report" in t and pkg.run_flag(r, "b") for r, t in runs), "plain text became bold")

for items, fmt in ((("Alpha", "Beta", "Gamma"), "bullet"), (("First", "Second", "Third"), "decimal")):
    ps = [by_text.get(t) for t in items]
    if ck.require(all(p is not None for p in ps), f"list items {items} missing"):
        fmts = [pkg.num_format(p) for p in ps]
        ck.require(all(f == fmt for f in fmts), f"list {items} has formats {fmts}, expected {fmt}")
        ck.require(len({pkg.num_pr(p)[0] for p in ps if pkg.num_pr(p)}) == 1, f"list {items} is not one list")

tables = list(pkg.body.iter(W + "tbl"))
tbl = next((t for t in tables if len(t.findall(W + "tr")) == 5
            and all(len(tr.findall(W + "tc")) == 4 for tr in t.findall(W + "tr"))), None)
if ck.require(tbl is not None, "no 5 x 4 table"):
    head = tbl.findall(W + "tr")[0].findall(W + "tc")
    ck.require([norm(pkg.text(c)) for c in head] == ["Region", "Q1", "Q2", "Q3"], "header row texts differ")
    fills = [c.find(f"{W}tcPr/{W}shd") for c in head]
    ok = all(s is not None and (s.get(W + "fill") or "auto").upper() not in ("AUTO", "FFFFFF") for s in fills)
    ck.require(ok, "header row cells are not shaded")
    ck.require(norm(pkg.text(tbl.findall(W + "tr")[4])).startswith("West"), "last row is not West")

pics = picture_targets(pkg, pkg.body)
ok = [p for p in pics if p[1] and png_size(pkg.read(p[1])) == (400, 300)]
ck.require(ok, "no PNG picture (400 x 300 px) in the body")

footer = pkg.story_for(sects, len(sects) - 1, "footer") if sects else None
if ck.require(footer, "the section has no default footer"):
    froot = pkg.xml(footer)
    fields = [f.split()[0].upper() for f in pkg.fields(froot) if f]
    ck.require("PAGE" in fields and "NUMPAGES" in fields and fields.index("PAGE") < fields.index("NUMPAGES"),
               f"footer fields are {fields}, expected PAGE then NUMPAGES")
    ftext = norm(pkg.text(froot))
    ck.require(ftext.startswith("Page") and " of " in f" {ftext} ", f"footer text {ftext!r} is not 'Page X of Y'")

body_fields = pkg.fields(pkg.body)
ck.require(any(f.upper().startswith("TOC") for f in body_fields), "no TOC field in the body")
entries = [p for p in paras if pkg.para_style_name(p).startswith("toc") and norm(pkg.text(p))]
ck.info(f"toc_entries={len(entries)}")
ck.done()
