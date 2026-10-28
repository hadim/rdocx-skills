"""docx: fixes the pinned integration build (upstream v0.16.4 plus tensorbee/rdocx#427 and #428) carries and upstream
`main` does not have yet, which the skills rely on. Each test names the pull request that carries it: delete the test
once upstream main has it. Each test builds its own input."""
import io
import re
import zipfile

import pytest
import rdocx


def rewritten(doc, edit):
    """doc saved, its parts passed to edit(parts) (a dict of name: bytes), and opened again."""
    archive = zipfile.ZipFile(io.BytesIO(doc.to_bytes()))
    parts = {name: archive.read(name) for name in archive.namelist()}
    edit(parts)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return rdocx.Document.from_bytes(out.getvalue())


def body_top(header_lines):
    """The y of the first body line on page 1, under a 2 cm top margin and a header of header_lines lines from
    1.27 cm."""
    doc = rdocx.Document()
    doc.update_section(0, margin_top=rdocx.Cm(2.0), header_distance=rdocx.Cm(1.27))
    for i in range(header_lines):
        doc.section(0).header.add_paragraph(f"Header line {i + 1}")
    for i in range(40):
        doc.add_paragraph(f"Body line {i + 1}.")
    fragments = doc.layout()
    first = min(f.physical_page for f in fragments)
    return min(f.bounds.y for f in fragments if f.physical_page == first)


# carried by tensorbee/rdocx#428 (#424), delete once upstream main has it
def test_body_starts_below_a_header_taller_than_the_top_margin():
    """A header that runs below the top margin moves the body down, as in Word, so renders, PDFs and the page
    numbers of a rebuilt TOC match Word's. The body used to start at the top margin, under the header's lines
    (tensorbee/rdocx#424)."""
    one, five = body_top(1), body_top(5)
    assert one == pytest.approx(2.0 / 2.54 * 72, abs=0.5)
    assert five > one + 3 * 10


def toc_document(merged, instruction):
    """Two Heading 1 paragraphs, a table between them (its first column merged vertically over two rows when
    merged) and a TOC field whose instruction is replaced by instruction."""
    doc = rdocx.Document()
    doc.add_paragraph("Title page")
    doc.add_paragraph("First section")
    doc.paragraphs[-1].style = "Heading 1"
    doc.add_table(3, 2, style="Table Grid")
    for r in range(3):
        for c in range(2):
            doc.tables[-1].cell(r, c).text = f"cell {r}{c}"
    if merged:
        doc.tables[-1].set_cell_vertical_merge(0, 0, "restart")
        doc.tables[-1].set_cell_vertical_merge(1, 0, "continue")
    doc.add_paragraph("Second section")
    doc.paragraphs[-1].style = "Heading 1"
    doc.insert_toc(1, max_level=2)

    def edit(parts):
        xml, n = re.subn(r"(<w:instrText[^>]*>)\s*TOC[^<]*(</w:instrText>)",
                         lambda m: m.group(1) + instruction + m.group(2), parts["word/document.xml"].decode())
        assert n == 1
        parts["word/document.xml"] = xml.encode()

    return rewritten(doc, edit)


# carried by tensorbee/rdocx#428 (#426), delete once upstream main has it
@pytest.mark.parametrize("merged", [False, True])
def test_rebuild_toc_with_u_switch_beside_a_vertical_merge(merged):
    """A TOC whose instruction carries \\u, as Word writes it for many templates, rebuilds in a document whose table
    has a vertically merged cell (python-api.md, rebuild_toc). It used to raise RdocxError ("TOC outline source has no
    qualified current layout facts", tensorbee/rdocx#426)."""
    doc = toc_document(merged, ' TOC \\o "1-2" \\h \\u ')
    assert doc.rebuild_toc().entry_count == 2


# carried by tensorbee/rdocx#427 (#425), delete once upstream main has it
def test_compare_keeps_the_revisions_of_a_redline():
    """A second reviewer's edit of a redline compares against it: the first reviewer's revisions stay as they are and
    the edit adds revisions by the second (python-api.md, compare). It used to raise RdocxError ("document comparison
    requires inputs without existing modeled revisions", tensorbee/rdocx#425)."""
    v1 = rdocx.Document()
    v1.add_paragraph("The first sentence stays. The second sentence changes once.")
    v1.add_paragraph("Another paragraph changes later.")
    v1 = rdocx.Document.from_bytes(v1.to_bytes())
    v2 = rdocx.Document.from_bytes(v1.to_bytes())
    v2.paragraphs[0].replace_text("changes once", "changed once", expect=1)
    redline = rdocx.Document.from_bytes(v1.to_bytes())
    redline.compare(v2, "Reviewer A", "2026-10-27T12:00:00Z", granularity="word")
    redline = rdocx.Document.from_bytes(redline.to_bytes())
    first = sorted((r.author, r.kind) for r in redline.revisions)
    edited = rdocx.Document.from_bytes(redline.to_bytes())
    edited.paragraphs[1].replace_text("changes later", "changed later", expect=1)
    second = rdocx.Document.from_bytes(redline.to_bytes())
    second.compare(edited, "Reviewer B", "2026-10-27T13:00:00Z", granularity="word")
    authors = [r.author for r in second.revisions]
    assert authors.count("Reviewer A") == len(first) and authors.count("Reviewer B") >= 1
