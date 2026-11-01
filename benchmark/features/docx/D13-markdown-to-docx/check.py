"""D13 oracle: the Markdown structures became Word structures, with no Markdown syntax left in the text."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg, norm, q  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
paras = pkg.paragraphs()
texts = [pkg.text(p) for p in paras]
for text, level in (("Project notes", 1), ("Tasks", 2), ("Steps", 3), ("Results", 2), ("Code", 2)):
    ck.require(any(norm(pkg.text(p)) == text and pkg.para_style_name(p) == f"heading {level}" for p in paras),
               f"no heading {level} '{text}'")
intro = next((p for p in paras if pkg.text(p).startswith("Intro paragraph")), None)
if ck.require(intro is not None, "intro paragraph missing"):
    ck.require(norm(pkg.text(intro)) == "Intro paragraph with bold text and italic text and a project link.",
               f"intro reads {pkg.text(intro)!r}")
    runs = pkg.runs_with_text(intro)
    ck.require(any("bold text" in t and pkg.run_flag(r, "b") for r, t in runs), "'bold text' is not bold")
    ck.require(any("italic text" in t and pkg.run_flag(r, "i") for r, t in runs), "'italic text' is not italic")
    links = [h for h in intro.iter(W + "hyperlink") if h.get(q("r:id"))]
    rel = pkg.rels(pkg.main).get(links[0].get(q("r:id"))) if links else None
    ck.require(rel and rel[1] == "https://example.com/project" and pkg.text(links[0]) == "project link",
               "no hyperlink 'project link' to https://example.com/project")
    # a link must look like one (Hyperlink style or underline), as Word, pandoc and LibreOffice write it
    ck.require(links and all(r.find(f"{W}rPr/{W}rStyle") is not None or pkg.run_flag(r, "u")
                             for r in links[0].iter(W + "r")), "the link looks like plain text (no style, no underline)")
for items, fmt in ((("Collect the samples", "Label the boxes", "Ship to the lab"), "bullet"),
                   (("Open the kit", "Fill the form", "Seal the box"), "decimal")):
    ps = [next((p for p in paras if norm(pkg.text(p)) == t), None) for t in items]
    if ck.require(None not in ps, f"list items {items} missing"):
        ck.require([pkg.num_format(p) for p in ps] == [fmt] * 3, f"{items} are not a {fmt} list")
        ck.require(len({pkg.num_pr(p) for p in ps}) == 1, f"{items} are not one list")
tbl = next(iter(pkg.body.iter(W + "tbl")), None)
if ck.require(tbl is not None, "no table"):
    got = [[norm(pkg.text(c)) for c in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]
    ck.require(got == [["Site", "Samples", "Status"], ["North", "12", "done"], ["South", "9", "pending"]],
               f"table cells differ: {got}")
code_lines = ["def total(xs):", "    return sum(xs)"]
code = [p for p in paras if pkg.text(p).strip().startswith(("def total", "return sum"))]
joined = "\n".join(pkg.text(p) for p in code).replace("\r", "")
ck.require(code_lines[0] in joined and code_lines[1] in joined, f"code block lost its lines or indent: {joined!r}")
mono = all(any(f in pkg.run_fonts(r) for f in ("courier", "consolas", "mono", "menlo", "source code"))
           for p in code for r, t in pkg.runs_with_text(p) if t.strip())
ck.require(code and mono, "code block is not in a monospace font")
ck.require("Final paragraph." in [norm(t) for t in texts], "last paragraph missing")
left = [t for t in texts if any(s in t for s in ("**", "```", "|---", "](")) or t.lstrip().startswith(("# ", "- "))]
ck.require(not left, f"Markdown syntax left in the text: {left}")
ck.done()
