"""R02 input: two open comment threads as Word writes them (comments.xml, commentsExtended.xml with done="0",
commentsIds.xml, people.xml), each anchored on text split over several runs."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

PAYMENT = ("The Customer shall pay each invoice within thirty days after delivery of the goods, unless the "
           "parties agree otherwise in writing.")
INTEREST = "Interest on late payment accrues at the statutory rate from the first day of delay."
BODY = [("Title", "Service agreement"), ("Heading1", "Payment"), (None, PAYMENT), (None, INTEREST)]
THREADS = [  # (comment, paragraph text, anchor)
    ({"id": 0, "author": "Alex Author", "initials": "AA", "date": "2026-10-01T08:30:00Z",
      "paragraphs": ["Customers asked for 45 days."]}, PAYMENT, "within thirty days"),
    ({"id": 1, "author": "Sam Lee", "initials": "SL", "date": "2026-10-02T14:05:00Z",
      "paragraphs": ["Which statute?"]}, INTEREST, "statutory rate"),
]


def anchored(text, anchor, cid):
    j = text.index(anchor)
    return (b.frag(text[:j], cuts=(7,)) + b.comment_start(cid) + b.frag(anchor, cuts=(6, 11))
            + b.comment_end(cid) + b.frag(text[j + len(anchor):], cuts=(9,)))


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    parts = []
    for style, text in BODY:
        thread = next((th for th in THREADS if th[1] == text), None)
        parts.append(b.para(anchored(text, thread[2], thread[0]["id"]) if thread else b.frag(text), style=style))
    pkg.comments([th[0] for th in THREADS])
    return pkg.save(Path(out_dir) / "input.docx", "".join(parts) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
