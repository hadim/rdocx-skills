"""R12 input: seven headings over five pages. Two Heading 1 start a page (page break before); the others fall
where the text before them ends, each well inside its page (no heading within a paragraph and a half of a page
edge in rdocx or LibreOffice), so the expected page does not depend on the renderer's exact line breaks."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

HEADINGS = [  # (level, text, expected page, filler paragraphs after it, page break before)
    (1, "Introduction", 1, 3, False),
    (1, "Scope of work", 1, 9, False),
    (2, "Deliverables", 2, 1, False),
    (2, "Acceptance tests", 2, 2, False),
    (1, "Payment terms", 3, 11, True),
    (2, "Late payment", 4, 2, False),
    (3, "Interest rate", 4, 3, False),
    (1, "Termination", 5, 2, True),
]
EXPECTED = [(lvl, text, page) for lvl, text, page, _, _ in HEADINGS]


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    body = [b.para(b.frag("Service agreement"), style="Title")]
    seed = 100
    for lvl, text, _page, n, brk in HEADINGS:
        body.append(b.para(b.frag(text), style=f"Heading{lvl}", page_break_before=brk))
        for _ in range(n):
            seed += 1
            body.append(b.para(b.frag(b.lorem(seed), cuts=(35, 150))))
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
