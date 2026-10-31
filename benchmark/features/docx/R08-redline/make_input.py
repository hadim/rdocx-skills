"""R08 inputs: v1 and v2 of an agreement, each saved with its own run fragmentation. v2 changes words in two
paragraphs and a table cell, removes a paragraph and adds one."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

V1 = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within thirty days after delivery of the goods."),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    (None, "Paper invoices are accepted until the end of the year."),
    ("Heading1", "Notices"),
    (None, "Notices are sent by registered letter to the addresses below."),
    ("TABLE", [["Party", "Address"], ["Customer", "12 Harbour Road, Portsmouth"]]),
    (None, "End of agreement."),
]
V2 = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within forty-five days after receipt of a valid invoice."),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    (None, "Invoices are sent by email to the address given by the Customer."),
    ("Heading1", "Notices"),
    (None, "Notices are sent by email or by registered letter to the addresses below."),
    ("TABLE", [["Party", "Address"], ["Customer", "Unit 4, 12 Harbour Road, Portsmouth"]]),
    (None, "End of agreement."),
]
AUTHOR = "Dana Reviewer"
REMOVED = "Paper invoices are accepted until the end of the year."


def flat(version):
    """(style, text) of every paragraph, table cells flattened, as the oracle reads a body."""
    out = []
    for style, text in version:
        if style == "TABLE":
            out += [(None, cell) for row in text for cell in row]
        else:
            out.append((style, text))
    return out


def write(version, path, seed):
    b.reset_ids()
    b.IDS.n += seed
    pkg = b.Package()
    body = []
    for k, (style, text) in enumerate(version):
        if style == "TABLE":
            body.append(b.table(text, widths=[3000, 6360]))
        else:
            cuts = (len(text) // 3 + seed % 5, 2 * len(text) // 3) if len(text) > 30 else ()
            body.append(b.para(b.frag(text, cuts=cuts), style=style))
    return pkg.save(path, "".join(body) + b.sect_pr())


def build(out_dir):
    write(V1, Path(out_dir) / "v1.docx", 0)
    return write(V2, Path(out_dir) / "v2.docx", 3)


if __name__ == "__main__":
    print(build(sys.argv[1]))
