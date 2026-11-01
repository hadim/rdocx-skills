"""D12 oracle: the 'Body Note' style (based on Normal, Arial 10 pt, 6 pt / 12 pt) on both notes; the second list
restarts at 1 as Word reads it (another w:num with a start override, or another abstract list), the first
list unchanged."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
style = next((s for sid, s in pkg.styles.items() if pkg.style_name(sid) == "Body Note"), None)
if ck.require(style is not None and style.get(W + "type") == "paragraph", "no paragraph style named 'Body Note'"):
    sid = style.get(W + "styleId")
    based = style.find(W + "basedOn")
    ck.require(based is not None and pkg.style_name(based.get(W + "val")) == "Normal", "not based on Normal")
    fonts = style.find(f"{W}rPr/{W}rFonts")
    ck.require(fonts is not None and fonts.get(W + "ascii") == "Arial", "font is not Arial")
    sz = style.find(f"{W}rPr/{W}sz")
    ck.require(sz is not None and sz.get(W + "val") == "20", "size is not 10 pt")
    sp = style.find(f"{W}pPr/{W}spacing")
    ck.require(sp is not None and sp.get(W + "before") == "120" and sp.get(W + "after") == "240",
               "spacing is not 6 pt before, 12 pt after")
    notes = [p for p in pkg.paragraphs() if pkg.text(p).startswith("Note:")]
    ck.require(len(notes) == 2 and all(pkg.para_style(p) == sid for p in notes), "the notes are not in 'Body Note'")
    ck.require(all(any(t == "Note:" and pkg.run_flag(r, "b") for r, t in pkg.runs_with_text(p)) for p in notes),
               "'Note:' lost its bold")
first = [pkg.find_para(lambda t, x=x: t == x) for x in ("Open the valve.", "Check the pressure.", "Close the valve.")]
second = [pkg.find_para(lambda t, x=x: t == x) for x in ("Open the bypass.", "Check the flow.", "Close the bypass.")]
if ck.require(None not in first + second, "list items missing"):
    a = {pkg.num_pr(p) for p in first}
    b = {pkg.num_pr(p) for p in second}
    ck.require(a == {("10", 0)}, f"the first list changed: {a}")
    ck.require(len(b) == 1 and None not in b and b != a, f"the second list is not one other list: {b}")
    ck.require(all(pkg.num_format(p) == "decimal" for p in second), "the second list is not decimal")
    if len(b) == 1 and None not in b and b != a:
        nid = next(iter(b))[0]
        num = pkg.num(nid)
        override = num.find(f"{W}lvlOverride[@{W}ilvl='0']/{W}startOverride") if num is not None else None
        other_abstract = pkg.abstract_of(nid) is not pkg.abstract_of("10")
        ck.require((override is not None and override.get(W + "val") == "1") or other_abstract,
                   "the second list's w:num neither overrides its start nor uses another abstract list: Word continues it")
        ck.info(f"restart by {'startOverride' if override is not None else 'another abstract list'}")
ck.done()
