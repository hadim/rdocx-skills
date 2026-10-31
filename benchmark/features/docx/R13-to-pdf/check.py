"""R13 oracle: OUT/document.pdf and OUT/tracked.pdf each have the document's four pages, with text or graphics
on them, and differ in their page content (the tracked view draws the revisions)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import PAGES  # noqa: E402


def main(out):
    doc, trk = Path(out) / "document.pdf", Path(out) / "tracked.pdf"
    a = o.pdf_info(doc) if doc.exists() else (0, False, 0)
    b = o.pdf_info(trk) if trk.exists() else (0, False, 0)
    o.report({
        "document_pages": a[0] == PAGES,
        "document_not_blank": a[1],
        "tracked_pages": b[0] == PAGES,
        "tracked_not_blank": b[1],
        "views_differ": a[2] != b[2],
    }, {"document": a[:2], "tracked": b[:2]})


o.guard(lambda: main(sys.argv[1]))
