"""R04 input: an agreement with a bold run-in lead, the word to change inside text split over runs, a table and
an earlier comment."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

LEAD = "Payment terms. "
PAYMENT = ("The Customer shall pay each invoice within thirty days after delivery of the goods, unless the "
           "parties agree otherwise in writing.")
BODY = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, LEAD + PAYMENT),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    ("Heading1", "Notices"),
    (None, "Notices are sent by registered letter to the addresses below."),
    (None, "Party"), (None, "Address"), (None, "Customer"), (None, "12 Harbour Road, Portsmouth"),
    (None, ""),
]
OLD, NEW = "thirty", "forty-five"
ACCEPTED = [(s, t.replace(OLD, NEW)) for s, t in BODY]


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    k = PAYMENT.index(OLD)
    body = [
        b.para(b.frag("Service agreement"), style="Title"),
        b.para(b.frag("Payment"), style="Heading1"),
        b.para(b.run(LEAD, b.rpr(bold=True)) + b.frag(PAYMENT, cuts=(k - 3, k + 3, k + 11), spell=((k, k + 6),))),
        b.para(b.frag(BODY[3][1], cuts=(30,))),
        b.para(b.frag("Notices"), style="Heading1"),
        b.para(b.frag(BODY[5][1], cuts=(12, 40))),
        b.table([["Party", "Address"], ["Customer", "12 Harbour Road, Portsmouth"]], widths=[3000, 6360]),
        b.para(""),
    ]
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
