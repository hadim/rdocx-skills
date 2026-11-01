"""R07 input: tracked changes by two authors, of each kind: inline insertions and deletions, an inserted and a
deleted paragraph, formatting changes, revisions in a table cell."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

A, D = "Alex Author", "Dana Reviewer"
TA, TD = "2026-10-05T10:00:00Z", "2026-10-06T09:15:00Z"
REJECT = D
# accepted view once Dana's changes are rejected (Alex's still pending, shown accepted)
EXPECTED_ACCEPTED = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within thirty days after delivery and receipt of a valid invoice."),
    (None, "Paper invoices are no longer accepted."),
    (None, "Interest on late payment accrues at the rate from the first day of delay."),
    ("Heading1", "Notices"),
    (None, "Notices are binding once received at the addresses below."),
    (None, "Party"), (None, "Address"), (None, "Customer"), (None, "12 Harbour Road, Portsmouth"),
    (None, "End of agreement."),
]
ORIGINAL = [  # every change rejected
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within thirty days after delivery."),
    (None, "Paper invoices are no longer accepted."),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    ("Heading1", "Notices"),
    (None, "Notices are binding once received at the addresses below."),
    (None, "Party"), (None, "Address"), (None, "Customer"), (None, "12 Harbour Road, Old Mill, Portsmouth"),
    (None, "End of agreement."),
]
ALEX_REVISIONS = [("ins", " and receipt of a valid invoice"), ("del", "statutory "), ("rPrChange", " once received"),
                  ("del", "Old Mill, ")]


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    body = [
        b.para(b.frag("Service agreement"), style="Title"),
        b.para(b.frag("Payment"), style="Heading1"),
        b.para(b.frag("The Customer shall pay each invoice within ", cuts=(13,)) + b.dele("thirty", D, TD)
               + b.ins(b.run("forty-five"), D, TD) + b.frag(" days after delivery")
               + b.ins(b.run(" and receipt of a valid invoice"), A, TA) + b.frag(".")),
        b.para(b.ins(b.frag("Invoices are sent by email.", cuts=(12,)), D, TD), mark_rpr=b.mark("ins", D, TD)),
        b.para(b.dele("Paper invoices are no longer accepted.", D, TD), mark_rpr=b.mark("del", D, TD)),
        b.para(b.frag("Interest on late payment accrues at the ", cuts=(9,)) + b.dele("statutory ", A, TA)
               + b.frag("rate from the first day of delay.")),
        b.para(b.frag("Notices"), style="Heading1"),
        b.para(b.run("Notices are binding", b.rpr_change("<w:b/><w:bCs/>", "", D, TD))
               + b.run(" once received", b.rpr_change("<w:i/><w:iCs/>", "", A, TA))
               + b.frag(" at the addresses below.")),
        b.table([["Party", "Address"],
                 ["Customer", b.para(b.frag("12 Harbour Road, ") + b.dele("Old Mill, ", A, TA)
                                     + b.ins(b.run("Unit 4, "), D, TD) + b.frag("Portsmouth"))]],
                widths=[3000, 6360]),
        b.para(b.frag("End of agreement.")),
    ]
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
