"""D11 oracle: one footnote with the text, referenced right after the sentence, as a superscript mark; the
paragraph's text and formatting unchanged."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
SENT = "We visited each region. The survey covered 120 sites. Results follow below."
p = pkg.find_para(lambda t: t.startswith("We visited"))
part = pkg.part_by_type("footnotes")
if ck.require(p is not None and part, "paragraph or footnotes part missing"):
    ck.require(pkg.text(p) == SENT, f"paragraph reads {pkg.text(p)!r}")
    refs = list(p.iter(W + "footnoteReference"))
    if ck.require(len(refs) == 1, f"{len(refs)} footnote references in the paragraph, expected 1"):
        ref = refs[0]
        before, after, seen = [], [], False
        for r, t in pkg.runs_with_text(p):
            if ref in r.iter(W + "footnoteReference"):
                seen = True
                continue
            (after if seen else before).append(t)
        ck.require("".join(before).endswith("The survey covered 120 sites.") and "".join(after).startswith(" Results"),
                   f"the mark sits after {''.join(before)[-30:]!r}")
        run = ref.getparent()
        va = run.find(f"{W}rPr/{W}vertAlign")
        rs = run.find(f"{W}rPr/{W}rStyle")
        styled = rs is not None and any(s.find(f"{W}rPr/{W}vertAlign") is not None and
                                        s.find(f"{W}rPr/{W}vertAlign").get(W + "val") == "superscript"
                                        for s in pkg.style_chain(rs.get(W + "val")))
        ck.require((va is not None and va.get(W + "val") == "superscript") or styled, "the mark is not superscript")
        fid = ref.get(W + "id")
        notes = [n for n in pkg.xml(part).iter(W + "footnote") if n.get(W + "id") == fid]
        if ck.require(len(notes) == 1, f"no footnote with id {fid}"):
            ck.require(pkg.text(notes[0]).strip() == "Source: field survey 2026.", f"note reads {pkg.text(notes[0])!r}")
            ck.require(any(True for _ in notes[0].iter(W + "footnoteRef")), "the note shows no number (w:footnoteRef)")
    runs = pkg.runs_with_text(p)
    ck.require(any(t == "120" and pkg.run_flag(r, "b") for r, t in runs), "'120' lost its bold")
    ck.require(any("region" in t and pkg.run_flag(r, "i") for r, t in runs), "'region' lost its italic")
    types = {n.get(W + "type") for n in pkg.xml(part).iter(W + "footnote")}
    ck.info(f"separator notes present: {'separator' in types and 'continuationSeparator' in types}")
ck.done()
