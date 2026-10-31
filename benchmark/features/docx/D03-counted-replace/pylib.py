import sys

from docx import Document

src, out, expected = sys.argv[1:4]
OLD, NEW = "Service Level Agreement", "Support Contract"
doc = Document(src)


def paragraphs():
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in doc.sections:
        for story in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            if not story.is_linked_to_previous:
                yield from story.paragraphs


def replace(p):
    count = 0
    while True:
        runs = p.runs
        text = "".join(r.text for r in runs)
        start = text.find(OLD)
        if start < 0:
            return count
        pos, first = 0, True
        for r in runs:
            a, b = pos, pos + len(r.text)
            pos = b
            if b > start and a < start + len(OLD):
                lo, hi = max(start, a) - a, min(start + len(OLD), b) - a
                r.text = r.text[:lo] + (NEW if first else "") + r.text[hi:]
                first = False
        count += 1


found = sum("".join(r.text for r in p.runs).count(OLD) for p in paragraphs())
if found != int(expected):
    sys.exit(f"refused: expected {expected} occurrence(s) of {OLD!r}, found {found}; nothing written")
for p in paragraphs():
    replace(p)
doc.save(out)
