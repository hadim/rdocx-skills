"""R03 input: four comments as Word writes them: a two-paragraph comment on text split over runs, a reply to it
(threaded in commentsExtended, its markers nested in its parent's), a comment whose range spans two paragraphs,
and a comment on part of a table cell."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

PAYMENT = ("The Customer shall pay each invoice within thirty days after delivery of the goods, unless the "
           "parties agree otherwise in writing.")
INTEREST = "Interest on late payment accrues at the statutory rate from the first day of delay."
DELIVERY = "Delivery takes place at the Customer's site during working hours."
COMMENTS = [
    {"id": 0, "author": "Alex Author", "initials": "AA", "date": "2026-10-01T08:30:00Z",
     "paragraphs": ["Customers asked for 45 days.", "See the email of 28 September."]},
    {"id": 1, "author": "Dana Reviewer", "initials": "DR", "date": "2026-10-03T16:45:00Z",
     "paragraphs": ["Agreed, let us discuss it on Monday."], "parent": 0},
    {"id": 2, "author": "Sam Lee", "initials": "SL", "date": "2026-10-02T14:05:00Z",
     "paragraphs": ["Does delay start before delivery?"]},
    {"id": 3, "author": "Alex Author", "initials": "AA", "date": "2026-10-04T11:20:00Z",
     "paragraphs": ["Net 30 contradicts clause 1."]},
]
EXPECTED = {  # id: (anchored text, parent id)
    0: ("within thirty days", None),
    1: ("within thirty days", 0),
    2: ("from the first day of delay.\nDelivery takes place", None),
    3: ("net 30", None),
}


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    i, j = PAYMENT.index("within thirty days"), INTEREST.index("from the first")
    k = len("Delivery takes place")
    body = [
        b.para(b.frag("Service agreement"), style="Title"),
        b.para(b.frag("Payment"), style="Heading1"),
        b.para(b.frag(PAYMENT[:i], cuts=(9,)) + b.comment_start(0) + b.comment_start(1)
               + b.frag("within thirty days", cuts=(4, 12)) + b.comment_end(0) + b.comment_end(1)
               + b.frag(PAYMENT[i + 18:], cuts=(20,))),  # a reply's markers as Word writes them
        b.para(b.frag(INTEREST[:j], cuts=(12,)) + b.comment_start(2) + b.frag(INTEREST[j:], cuts=(9,))),
        b.para(b.frag(DELIVERY[:k], cuts=(5,)) + b.comment_end(2) + b.frag(DELIVERY[k:])),
        b.table([["Item", "Terms"],
                 ["Invoices", b.para(b.frag("Payable ") + b.comment_start(3) + b.frag("net 30", cuts=(2,))
                                     + b.comment_end(3) + b.frag(" from receipt"))],
                 ["Credit notes", "Issued within ten days"]], widths=[3000, 6360]),
        b.para(""),
    ]
    pkg.comments(COMMENTS)
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
