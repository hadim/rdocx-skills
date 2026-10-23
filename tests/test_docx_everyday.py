"""docx: what the pinned build lacks of the everyday surface the skills cite, each as a strict xfail
(tests/gaps.py). Each test builds its own input."""
import io
import re
import zipfile

import pytest
import rdocx

from builders import W, noise_png
from conftest import digest


@pytest.mark.gap("docx-default-line-spacing-multiple")
def test_default_line_spacing_takes_a_multiple():
    """set_default_paragraph_format takes a length (an exact pitch) but no multiple, the rule of Word's own default
    (1.08 lines); reading a multiple works."""
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


@pytest.mark.gap("docx-insert-document-reused-picture-swapped")
def test_insert_document_keeps_each_inserted_picture():
    """insert_document reuses the target's picture parts that hold the bytes of the inserted pictures, but pairs them
    in the wrong order: inserted red then blue, into a target that holds blue then red, show blue then red."""
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


@pytest.mark.gap("docx-comments-range-across-cells")
@pytest.mark.parametrize("start", ['<w:commentRangeStart w:id="0"/>', GOOG_START], ids=["plain", "goog_rdk"])
def test_comments_lists_a_range_across_two_cells(start):
    """A comment whose range starts in one cell and ends in the next one, as Google Docs writes a comment on a
    selection across cells (its start in a goog_rdk content control or not): Document.comments raises RdocxError
    ("cannot be projected faithfully"), remove_comment works."""
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
    assert [(c.id, c.text) for c in doc.comments] == [(0, "a remark")]


@pytest.mark.gap("docx-part-xml-without-relationships")
def test_part_xml_reads_relationships_and_content_types():
    """part_names() and part_xml() leave out the relationship parts and [Content_Types].xml (RdocxError "part ... is
    missing"): a tool that resolves a relationship id or checks a content type still unzips to_bytes()."""
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


@pytest.mark.gap("docx-insert-document-duplicate-style-ids")
def test_insert_document_maps_a_style_id_defined_twice():
    """The target and the inserted document carry the same styles part, which defines Heading1 and Title twice:
    insert_document adds Heading1Merge1 and TitleMerge1 and gives them to the inserted paragraphs."""
    target = heading_and_title("target")
    before = style_ids(target)
    target.insert_document(heading_and_title("inserted"), at=len(target.paragraphs))
    assert [i for i in style_ids(target) if i not in before] == []
    assert [p.style for p in target.paragraphs] == ["Heading1", "Title", "Heading1", "Title"]


@pytest.mark.gap("docx-insert-document-new-document-refused")
def test_insert_document_takes_a_new_document():
    """A target whose Normal style differs from rdocx's own, as in every document Word, Google Docs or python-docx
    writes: insert_document of a new Document() raises RdocxError ("invalid style graph: style type 'paragraph' has
    more than one default"), whatever the policy."""
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
