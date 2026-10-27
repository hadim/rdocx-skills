"""docx: fixes the pinned integration build (integration/open-prs-2026-10-27) carries and upstream `main` does not
have yet, which the skills rely on. Each test names the pull request that carries it: delete the test once upstream
main has it. Each test builds its own input."""
import io
import re
import zipfile

import pytest
import rdocx

from builders import W, noise_png
from conftest import digest


# carried by tensorbee/rdocx#414 (#402), delete once upstream main has it
def test_default_line_spacing_takes_a_multiple():
    """set_default_paragraph_format takes a number of lines, the rule of Word's own default (python-api.md, document
    defaults). It used to take lengths only and raise TypeError for 1.15 (tensorbee/rdocx#402)."""
    doc = rdocx.Document()
    doc.set_default_paragraph_format(line_spacing=1.15)
    assert doc.default_paragraph_format.line_spacing == pytest.approx(1.15)
    assert 'w:line="276" w:lineRule="auto"' in doc.part_xml("word/styles.xml")


def with_pictures(*pictures):
    """A document whose body holds one picture per paragraph, in that order."""
    doc = rdocx.Document()
    for data in pictures:
        doc.add_picture(data, "picture.png", rdocx.Cm(1), rdocx.Cm(1))
    return rdocx.Document.from_bytes(doc.to_bytes())


# carried by tensorbee/rdocx#411 (#404), delete once upstream main has it
def test_insert_document_keeps_each_inserted_picture():
    """insert_document into a target that holds the same pictures: each inserted picture keeps its own bytes
    (python-api.md, insert_document). Inserted red then blue, into a target that holds blue then red, used to show
    blue then red (tensorbee/rdocx#404)."""
    red, blue = noise_png(8, 1), noise_png(8, 2)
    target = with_pictures(blue, red)
    target.insert_document(with_pictures(red, blue), at=2)
    assert [digest(p.bytes) for p in target.pictures()] == [digest(d) for d in (blue, red, red, blue)]


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


GOOG_START = ('<w:sdt><w:sdtPr><w:id w:val="-346012225"/><w:tag w:val="goog_rdk_0"/></w:sdtPr><w:sdtContent>'
              '<w:commentRangeStart w:id="0"/></w:sdtContent></w:sdt>')


# carried by tensorbee/rdocx#412 (#405), delete once upstream main has it
@pytest.mark.parametrize("start", ['<w:commentRangeStart w:id="0"/>', GOOG_START], ids=["plain", "goog_rdk"])
def test_comments_lists_a_range_across_two_cells(start):
    """A comment whose range starts in one cell and ends in the next one, as Google Docs writes a comment on a
    selection across cells (its start in a goog_rdk content control or not): Document.comments lists it with its
    anchor text across the cells (SKILL.md, comments). It used to raise RdocxError (tensorbee/rdocx#405)."""
    doc = rdocx.Document()
    t = doc.add_table(1, 2)
    t.cell(0, 0).paragraphs[0].add_run("first cell")
    t.cell(0, 1).paragraphs[0].add_run("second cell")

    def edit(parts):
        xml = parts["word/document.xml"].decode()
        first = xml.index("<w:r>", xml.index("<w:tc>"))
        xml = xml[:first] + start + xml[first:]
        end = xml.index("</w:p>", xml.index("second cell"))
        xml = xml[:end] + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' + xml[end:]
        parts["word/document.xml"] = xml.encode()
        parts["word/comments.xml"] = (f'<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="A"><w:p><w:r>'
                                      '<w:t>a remark</w:t></w:r></w:p></w:comment></w:comments>').encode()
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(b"</Types>", (
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.comments+xml"/></Types>').encode())
        parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(b"</Relationships>", (
            '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'comments" Target="comments.xml"/></Relationships>').encode())

    doc = rewritten(doc, edit)
    assert [(c.id, c.text, c.anchor_text, c.anchor_issue) for c in doc.comments] == [
        (0, "a remark", "first cell\nsecond cell", None)]


# carried by tensorbee/rdocx#409 (#406), delete once upstream main has it
def test_part_xml_reads_relationships_and_content_types():
    """part_names() and part_xml() list and read the relationship parts and [Content_Types].xml as to_bytes() writes
    them (python-api.md, Raw XML). They used to leave them out (tensorbee/rdocx#406)."""
    doc = rdocx.Document()
    doc.add_paragraph("x")
    saved = zipfile.ZipFile(io.BytesIO(doc.to_bytes()))
    for name in ("/[Content_Types].xml", "/_rels/.rels", "/word/_rels/document.xml.rels"):
        assert name in doc.part_names()
        assert digest(doc.part_xml(name).encode()) == digest(saved.read(name[1:]))


SECOND_DEFINITION = ('<w:style w:type="paragraph" w:styleId="{id}"><w:name w:val="{name}"/><w:basedOn w:val="Normal"/>'
                     '<w:rPr><w:color w:val="C00000"/></w:rPr></w:style>')


def heading_and_title(text):
    """A document of one Heading 1 and one Title paragraph whose styles part defines Heading1 and Title a second time,
    with another colour, as Google Docs writes some styles parts."""
    doc = rdocx.Document()
    doc.add_heading(text, level=1)
    doc.add_heading(text, level=0)
    extra = "".join(SECOND_DEFINITION.format(id=i, name=n) for i, n in (("Heading1", "heading 1"), ("Title", "Title")))

    def edit(parts):
        parts["word/styles.xml"] = parts["word/styles.xml"].replace(b"</w:styles>", extra.encode() + b"</w:styles>")
    return rewritten(doc, edit)


def style_ids(doc):
    return re.findall(r'w:styleId="([^"]+)"', doc.part_xml("word/styles.xml"))


# carried by tensorbee/rdocx#410 (#408), delete once upstream main has it
def test_insert_document_maps_a_style_id_defined_twice():
    """The target and the inserted document carry the same styles part, which defines Heading1 and Title twice:
    insert_document adds no style and the inserted paragraphs keep Heading1 and Title. It used to add Heading1Merge1
    and TitleMerge1 (tensorbee/rdocx#408)."""
    target = heading_and_title("target")
    before = style_ids(target)
    target.insert_document(heading_and_title("inserted"), at=len(target.paragraphs))
    assert [i for i in style_ids(target) if i not in before] == []
    assert [p.style for p in target.paragraphs] == ["Heading1", "Title", "Heading1", "Title"]


# carried by tensorbee/rdocx#410, delete once upstream main has it
def test_insert_document_takes_a_new_document():
    """A target whose Normal style differs from rdocx's own, as in every document Word, Google Docs or python-docx
    writes, takes a new Document() through insert_document (python-api.md, insert_document). It used to raise
    RdocxError ("invalid style graph: style type 'paragraph' has more than one default")."""
    doc = rdocx.Document()
    doc.add_paragraph("target")

    def edit(parts):
        xml = parts["word/styles.xml"].decode()
        at = xml.index(">", xml.index('w:styleId="Normal"')) + 1
        parts["word/styles.xml"] = (xml[:at] + '<w:rPr><w:sz w:val="22"/></w:rPr>' + xml[at:]).encode()

    target, source = rewritten(doc, edit), rdocx.Document()
    source.add_paragraph("inserted")
    target.insert_document(source, at=1)
    assert [p.text for p in target.paragraphs] == ["target", "inserted"]


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
