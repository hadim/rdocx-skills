"""D03 oracle: 5 replacements in the right places, nothing else changed, formatting around kept."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
NEW = "Support Contract"


def chars(p):
    """[(char, bold, italic, colour)] of a paragraph's runs."""
    out = []
    for r, t in pkg.runs_with_text(p):
        c = r.find(f"{W}rPr/{W}color")
        col = c.get(W + "val") if c is not None else None
        out += [(ch, pkg.run_flag(r, "b"), pkg.run_flag(r, "i"), col) for ch in t]
    return out


def span(p, sub):
    cs = chars(p)
    text = "".join(c[0] for c in cs)
    k = text.find(sub)
    return cs[k:k + len(sub)] if k >= 0 else None


expected_body = ["Contract summary", "Scope: The Support Contract applies to all customers.",
                 "Under the Support Contract, response times are guaranteed.",
                 "Support Contract terms are reviewed yearly.", "The service level is measured monthly.",
                 "Document", "Owner", "Support Contract", "Legal", "End of summary."]
body = [p for p in pkg.paragraphs() if pkg.text(p)]
texts = [pkg.text(p) for p in body]
ck.require(texts == expected_body, f"body text differs: {texts}")
header = pkg.story_for(pkg.sect_prs(), 0, "header")
htext = pkg.text(pkg.xml(header)) if header else ""
ck.require(htext == "Support Contract - draft for review", f"header text is {htext!r}")
if texts == expected_body:
    p1, p2, p3 = body[1], body[2], body[3]
    s = span(p1, "Scope: ")
    ck.require(all(c[1] for c in s), "'Scope: ' lost its bold")
    ck.require(all(c[1] for c in span(p1, "all customers")), "'all customers' lost its bold")
    ck.require(not any(c[1] for c in span(p1, "The " + NEW + " applies to ")), "plain text around became bold")
    ck.require(all(c[2] and not c[1] for c in span(p2, "Under the " + NEW + ", response")),
               "italic sentence lost its italic (or became bold)")
    ck.require(all(c[1] and c[2] for c in span(p2, "guaranteed")), "'guaranteed' lost bold italic")
    ck.require(not any(c[1] for c in span(p3, " terms are reviewed yearly.")), "text after the bold word became bold")
    bold_new = [c[1] for c in span(p3, NEW)]
    ck.info(f"mixed-format occurrence: new text bold={all(bold_new)} (first matched run was bold)")
    ck.require(all((c[3] or "").upper() == "C00000" for c in span(body[7], NEW)), "table cell text lost its red")
ck.done()
