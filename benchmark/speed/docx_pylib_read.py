"""All text of a .docx with python-docx: body paragraphs and tables (cells once, nested tables too), in order,
then headers and footers. Prints plain text."""
import sys

from docx import Document
from docx.table import Table


def table_lines(table, seen):
    for row in table.rows:
        for cell in row.cells:
            if cell._tc in seen:  # python-docx repeats a merged cell in each grid column it spans
                continue
            seen.add(cell._tc)
            for block in cell.iter_inner_content():
                if isinstance(block, Table):
                    yield from table_lines(block, seen)
                else:
                    yield block.text


doc = Document(sys.argv[1])
out, seen = [], set()
for block in doc.iter_inner_content():
    if isinstance(block, Table):
        out.extend(table_lines(block, seen))
    else:
        out.append(block.text)
for section in doc.sections:
    for story in (section.header, section.footer):
        if not story.is_linked_to_previous:
            out.extend(p.text for p in story.paragraphs)
sys.stdout.write("\n".join(out) + "\n")
