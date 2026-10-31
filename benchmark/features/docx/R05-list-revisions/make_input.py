"""R05 input: tracked changes by two authors as Word writes them: an insertion and a deletion in the body, a
formatting change (w:rPrChange), an insertion in a table cell and a deletion in a footnote."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

A, D = "Alex Author", "Dana Reviewer"
EXPECTED = [  # (type, author, date, text)
    ("insert", A, "2026-10-05T10:00:00Z", " and on receipt of a valid invoice"),
    ("delete", D, "2026-10-06T09:15:00Z", "statutory "),
    ("format", A, "2026-10-05T10:05:00Z", "Notices are binding"),
    ("insert", D, "2026-10-06T09:20:00Z", "Unit 4, "),
    ("delete", A, "2026-10-05T10:10:00Z", " as amended"),
]


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    e = {x[3]: x for x in EXPECTED}
    ins = lambda text: b.ins(b.run(text), e[text][1], e[text][2])  # noqa: E731
    dele = lambda text: b.dele(text, e[text][1], e[text][2])  # noqa: E731
    bold = b.rpr_change("<w:b/><w:bCs/>", "", A, e["Notices are binding"][2])
    body = [
        b.para(b.frag("Service agreement"), style="Title"),
        b.para(b.frag("Payment"), style="Heading1"),
        b.para(b.frag("The Customer shall pay each invoice within thirty days after delivery", cuts=(25,))
               + ins(" and on receipt of a valid invoice") + b.frag(".")),
        b.para(b.frag("Interest on late payment accrues at the ", cuts=(9,)) + dele("statutory ")
               + b.frag("rate from the first day of delay.") + b.footnote_ref(1)),
        b.para(b.frag("Notices"), style="Heading1"),
        b.para(b.run("Notices are binding", bold) + b.frag(" once received at the addresses below.")),
        b.table([["Party", "Address"],
                 ["Customer", b.para(b.frag("12 Harbour Road, ") + ins("Unit 4, ") + b.frag("Portsmouth"))]],
                widths=[3000, 6360]),
        b.para(""),
    ]
    pkg.notes({1: b.frag("Late Payment of Commercial Debts Act 1998") + dele(" as amended") + b.frag(".")})
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
