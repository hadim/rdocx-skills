"""R06 oracle: no revision markup in any story; the body equals the accepted text paragraph by paragraph (a
whole deleted paragraph leaves nothing, the one before the spacer leaves only the spacer); the formatting change
kept (bold); the footnote accepted."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import ACCEPTED, BOLD_RUN, FOOTNOTE  # noqa: E402


def main(out):
    d = o.Docx(out)
    got = o.body(d)
    fn_part = d.part_by_type("footnotes", d.main_part())
    fns = d.xml(fn_part).findall("w:footnote", o.NS) if fn_part else []
    fn_text = [o.run_text(p).strip() for f in fns if f.get(o.q("w:type")) is None for p in f.findall("w:p", o.NS)]
    bold = []
    for p in d.document.iter(o.q("w:p")):
        if o.run_text(p).startswith(BOLD_RUN):
            pos = 0
            for r in p.findall("w:r", o.NS):
                if pos < len(BOLD_RUN) and o.run_text(r):
                    bold.append(bool(o.effective_run(d, p, r).get("b")))
                pos += len(o.run_text(r))
    empties = sum(1 for _, t in got if t == "")
    o.report({
        "no_revisions_left": not o.all_revisions(d),
        "accepted_text": [t for _, t in got] == [t for _, t in ACCEPTED],
        "styles_kept": got == o.norm(ACCEPTED),
        "no_leftover_empty_paragraph": empties == sum(1 for _, t in ACCEPTED if t == ""),
        "no_empty_list_item": not [p for p in d.document.iter(o.q("w:p")) if not o.run_text(p)
                                   and o.numbering_format(d, p) is not None],
        "format_change_kept": bool(bold) and all(bold),
        "footnote_accepted": fn_text == [FOOTNOTE],
    }, {"body": got, "footnotes": fn_text})


o.guard(lambda: main(sys.argv[1]))
