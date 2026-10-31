"""R07 oracle: no revision by Dana Reviewer left; Alex Author's four revisions still there, unchanged; the
accepted view is the text with Dana's changes undone and Alex's applied, the rejected view the original; Dana's
bold undone, Alex's italic still tracked."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import ALEX_REVISIONS, EXPECTED_ACCEPTED, ORIGINAL, REJECT  # noqa: E402


def rev_text(el):
    if el.tag == o.q("w:rPrChange"):
        el = el.getparent().getparent()
    return "".join(t.text or "" for t in el.iter(o.q("w:t"), o.q("w:delText")))


def main(out):
    d = o.Docx(out)
    revs = [r for _, r in o.all_revisions(d)]
    alex = sorted((r.tag.split("}")[1], rev_text(r)) for r in revs if r.get(o.q("w:author")) == "Alex Author")
    fmt = {}
    for p in d.document.iter(o.q("w:p")):
        for r in p.findall("w:r", o.NS):
            txt = o.run_text(r)
            if txt in ("Notices are binding", " once received"):
                fmt[txt] = o.effective_run(d, p, r)
    o.report({
        "no_revision_by_rejected_author": not [r for r in revs if r.get(o.q("w:author")) == REJECT],
        "other_author_revisions_kept": alex == sorted(ALEX_REVISIONS),
        "accepted_view": o.body(d, "accepted") == o.norm(EXPECTED_ACCEPTED),
        "rejected_view": o.body(d, "rejected") == o.norm(ORIGINAL),
        "rejected_bold_undone": not fmt.get("Notices are binding", {"b": True}).get("b"),
        "kept_italic": bool(fmt.get(" once received", {}).get("i")),
    }, {"alex": alex, "accepted": o.body(d, "accepted")})


o.guard(lambda: main(sys.argv[1]))
