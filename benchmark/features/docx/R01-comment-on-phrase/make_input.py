"""R01 input: a short agreement whose payment paragraph holds the phrase split over four runs (rsid splits and
a spell-check mark), and an earlier comment by another author on another paragraph."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

PHRASE = "thirty days after delivery"
PAYMENT = ("The Customer shall pay each invoice within thirty days after delivery of the goods, unless the "
           "parties agree otherwise in writing.")
INTEREST = "Interest on late payment accrues at the statutory rate from the first day of delay."
OLD_ANCHOR = "late payment"
BODY = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, PAYMENT),
    (None, INTEREST),
    ("Heading1", "Delivery"),
    (None, "Delivery takes place at the Customer's site. The goods are inspected within five working days "
           "after delivery."),
]
OLD_COMMENT = {"id": 0, "author": "Alex Author", "initials": "AA", "date": "2026-10-01T08:30:00Z",
               "paragraphs": ["Is the statutory rate defined somewhere?"]}


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    k = PAYMENT.index(PHRASE)
    parts = []
    for style, text in BODY:
        if text == PAYMENT:
            content = b.frag(text, cuts=(k + 4, k + 14, k + 22), spell=((k + 14, k + 22),))
        elif text == INTEREST:
            j = text.index(OLD_ANCHOR)
            content = (b.frag(text[:j], cuts=(5,)) + b.comment_start(0) + b.frag(OLD_ANCHOR, cuts=(5,))
                       + b.comment_end(0) + b.frag(text[j + len(OLD_ANCHOR):]))
        else:
            content = b.frag(text, cuts=(len(text) // 2,) if len(text) > 30 else ())
        parts.append(b.para(content, style=style))
    pkg.comments([OLD_COMMENT])
    return pkg.save(Path(out_dir) / "input.docx", "".join(parts) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
