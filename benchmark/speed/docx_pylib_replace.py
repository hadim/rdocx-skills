"""Counted replacement with python-docx (the pylib stack): body paragraphs, table cells (each merged cell once,
nested tables too), headers and footers, across runs (the new text takes the first matched run); refuses,
writing nothing, when the count differs.
    python docx_pylib_replace.py IN OUT OLD NEW COUNT"""
import sys

from docx import Document

src, out, old, new, expected = sys.argv[1:6]
doc = Document(src)


def cell_paragraphs(table, seen):
    for row in table.rows:
        for cell in row.cells:
            if cell._tc in seen:  # a merged cell is repeated in every grid column it spans
                continue
            seen.add(cell._tc)
            yield from cell.paragraphs
            for inner in cell.tables:
                yield from cell_paragraphs(inner, seen)


def paragraphs():
    yield from doc.paragraphs
    seen = set()
    for table in doc.tables:
        yield from cell_paragraphs(table, seen)
    for section in doc.sections:
        for story in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            if not story.is_linked_to_previous:
                yield from story.paragraphs


def replace(p):
    count = 0
    while True:
        runs = p.runs
        text = "".join(r.text for r in runs)
        start = text.find(old)
        if start < 0:
            return count
        pos, first = 0, True
        for r in runs:
            a, b = pos, pos + len(r.text)
            pos = b
            if b > start and a < start + len(old):
                lo, hi = max(start, a) - a, min(start + len(old), b) - a
                r.text = r.text[:lo] + (new if first else "") + r.text[hi:]
                first = False
        count += 1


found = sum("".join(r.text for r in p.runs).count(old) for p in paragraphs())
if found != int(expected):
    sys.exit(f"refused: expected {expected} occurrence(s) of {old!r}, found {found}; nothing written")
for p in paragraphs():
    replace(p)
doc.save(out)
