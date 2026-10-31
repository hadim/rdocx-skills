"""R14 inputs: v1 and v2 of an agreement, each saved with its own run fragmentation (unchanged paragraphs split
differently into runs). v2 changes two paragraphs and a table cell, adds a paragraph and removes one."""
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
    ("Heading1", "Delivery"),
    (None, "Delivery takes place at the Customer's site during working hours."),
    ("TABLE", [["Item", "Terms"], ["Invoices", "Payable net 30 from receipt"], ["Credit notes", "Issued within ten days"]]),
    ("Heading1", "Notices"),
    (None, "Notices are sent by registered letter to the addresses below."),
    (None, "End of agreement."),
]
V2 = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within forty-five days after delivery of the goods."),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    ("Heading1", "Delivery"),
    (None, "Delivery takes place at the Customer's site during working hours."),
    (None, "The Supplier gives two working days' notice of each delivery."),
    ("TABLE", [["Item", "Terms"], ["Invoices", "Payable net 45 from receipt"], ["Credit notes", "Issued within ten days"]]),
    ("Heading1", "Notices"),
    (None, "Notices are sent by email or by registered letter to the addresses below."),
    (None, "End of agreement."),
]
EXPECTED = {
    "changed": [
        ["The Customer shall pay each invoice within thirty days after delivery of the goods.",
         "The Customer shall pay each invoice within forty-five days after delivery of the goods."],
        ["Payable net 30 from receipt", "Payable net 45 from receipt"],
        ["Notices are sent by registered letter to the addresses below.",
         "Notices are sent by email or by registered letter to the addresses below."],
    ],
    "added": ["The Supplier gives two working days' notice of each delivery."],
    "removed": ["Paper invoices are accepted until the end of the year."],
}


def write(version, path, seed):
    b.reset_ids()
    b.IDS.n += seed
    pkg = b.Package()
    body = []
    for style, text in version:
        if style == "TABLE":
            body.append(b.table(text, widths=[3000, 6360]))
        else:
            cuts = (len(text) // 4 + seed, len(text) // 2 + 2 * seed) if len(text) > 30 else ()
            body.append(b.para(b.frag(text, cuts=cuts), style=style))
    return pkg.save(path, "".join(body) + b.sect_pr())


def build(out_dir):
    write(V1, Path(out_dir) / "v1.docx", 0)
    return write(V2, Path(out_dir) / "v2.docx", 5)


if __name__ == "__main__":
    print(build(sys.argv[1]))
