"""D10 oracle: an external hyperlink and an internal one to a bookmark that sits in the target heading."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg, q  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
p = pkg.find_para(lambda t: t.startswith("Further reading:"))
if ck.require(p is not None, "no 'Further reading:' paragraph"):
    ck.require(pkg.text(p) == "Further reading: Project site and Appendix A.", f"paragraph reads {pkg.text(p)!r}")
    links = p.findall(W + "hyperlink")
    ext = [h for h in links if h.get(q("r:id"))]
    if ck.require(len(ext) == 1, f"{len(ext)} external links, expected 1"):
        rel = pkg.rels(pkg.main).get(ext[0].get(q("r:id")))
        ck.require(rel and rel[0] == "hyperlink" and rel[2] == "External" and rel[1] == "https://example.com/project",
                   f"external link relationship is {rel}")
        ck.require(pkg.text(ext[0]) == "Project site", f"external link text is {pkg.text(ext[0])!r}")
    internal = [h for h in links if h.get(W + "anchor")]
    if ck.require(len(internal) == 1, f"{len(internal)} internal links, expected 1"):
        name = internal[0].get(W + "anchor")
        ck.require(pkg.text(internal[0]) == "Appendix A", f"internal link text is {pkg.text(internal[0])!r}")
        heading = next((h for h in pkg.paragraphs() if pkg.text(h) == "Appendix A"
                        and pkg.para_style_name(h) == "heading 1"), None)
        if ck.require(heading is not None, "heading 'Appendix A' missing"):
            starts = [b for b in heading.iter(W + "bookmarkStart") if b.get(W + "name") == name]
            ck.require(starts, f"bookmark {name!r} is not in the heading")
            if starts:
                bid = starts[0].get(W + "id")
                ends = [e for e in pkg.body.iter(W + "bookmarkEnd") if e.get(W + "id") == bid]
                ck.require(len(ends) == 1, f"bookmark {name!r} has {len(ends)} ends")
        ids = [b.get(W + "id") for b in pkg.body.iter(W + "bookmarkStart")]
        ck.require(len(ids) == len(set(ids)), f"duplicate bookmark ids {ids}")
    styled = [h for h in links if any(r.find(f"{W}rPr/{W}rStyle") is not None or r.find(f"{W}rPr/{W}u") is not None
                                      for r in h.iter(W + "r"))]
    ck.require(len(styled) == len(links), f"only {len(styled)} of {len(links)} links look like links (style or underline)")
ck.done()
