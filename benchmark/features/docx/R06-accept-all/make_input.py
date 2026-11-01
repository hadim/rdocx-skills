"""R06 input: tracked changes of every common kind: inline insertion and deletion, an inserted paragraph, a
deleted whole paragraph, a deleted paragraph followed by an empty spacer paragraph, a formatting change, an
insertion and a deletion in a table cell, an insertion in a footnote."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

A, D = "Alex Author", "Dana Reviewer"
T1, T2 = "2026-10-05T10:00:00Z", "2026-10-06T09:15:00Z"
ACCEPTED = [
    ("Title", "Service agreement"),
    ("Heading1", "Payment"),
    (None, "The Customer shall pay each invoice within forty-five days after delivery."),
    (None, "Invoices are sent by email."),
    (None, "Interest on late payment accrues at the statutory rate from the first day of delay."),
    (None, "Accepted payment methods:"),
    ("ListParagraph", "Bank transfer"),
    ("ListParagraph", "Card"),
    (None, ""),
    ("Heading1", "Notices"),
    (None, "Notices are binding once received at the addresses below."),
    (None, "Party"), (None, "Address"), (None, "Customer"), (None, "12 Harbour Road, Unit 4, Portsmouth"),
    (None, "End of agreement."),
]
FOOTNOTE = "Late Payment of Commercial Debts Act 1998, as amended."
BOLD_RUN = "Notices are binding"


def deleted_paragraph(text, author, date, numbered=False):
    """A whole paragraph deleted as Word records it: every run in w:del and the paragraph mark deleted."""
    num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if numbered else ""
    return b.para(b.dele(text, author, date), style="ListParagraph" if numbered else None, ppr=num,
                  mark_rpr=b.mark("del", author, date))


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    body = [
        b.para(b.frag("Service agreement"), style="Title"),
        b.para(b.frag("Payment"), style="Heading1"),
        b.para(b.frag("The Customer shall pay each invoice within ", cuts=(13,)) + b.dele("thirty", A, T1)
               + b.ins(b.run("forty-five"), A, T1) + b.frag(" days after delivery.")),
        b.para(b.ins(b.frag("Invoices are sent by email.", cuts=(12,)), D, T2), mark_rpr=b.mark("ins", D, T2)),
        deleted_paragraph("Paper invoices are no longer accepted.", D, T2),
        b.para(b.frag("Interest on late payment accrues at the statutory rate from the first day of delay.",
                      cuts=(20,)) + b.footnote_ref(1)),
        b.para(b.frag("Accepted payment methods:")),
        b.numbered(b.frag("Bank transfer"), 1),
        deleted_paragraph("Direct debit", A, T1, numbered=True),  # a whole list item deleted
        b.numbered(b.frag("Card"), 1),
        deleted_paragraph("Cheque", A, T1, numbered=True),  # the last item deleted, then an empty spacer
        b.para(""),
        b.para(b.frag("Notices"), style="Heading1"),
        b.para(b.run(BOLD_RUN, b.rpr_change("<w:b/><w:bCs/>", "", D, T2))
               + b.frag(" once received at the addresses below.")),
        b.table([["Party", "Address"],
                 ["Customer", b.para(b.frag("12 Harbour Road, ") + b.dele("Old Mill, ", A, T1)
                                     + b.ins(b.run("Unit 4, "), D, T2) + b.frag("Portsmouth"))]],
                widths=[3000, 6360]),
        b.para(b.frag("End of agreement.")),
    ]
    pkg.numbering([(0, 1, "bullet", "\u2022", 720, 360)])
    pkg.notes({1: b.frag("Late Payment of Commercial Debts Act 1998") + b.ins(b.run(", as amended"), D, T2)
               + b.frag(".")})
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
