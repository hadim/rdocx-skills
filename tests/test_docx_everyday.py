"""docx: what the pinned build lacks of the everyday surface the skills cite, each as a strict xfail
(tests/gaps.py). Each test builds its own input."""
import pytest
import rdocx


@pytest.mark.gap("docx-custom-properties-not-listed")
def test_custom_properties_are_listed(tmp_path):
    """custom_properties takes item assignment, deletion and `in`, and get(name).value reads one property, but
    nothing lists the properties a document holds (no keys, items or iteration)."""
    doc = rdocx.Document()
    doc.custom_properties["Client"] = "Acme"
    doc.custom_properties["Revision"] = 3
    doc.custom_properties["Due"] = rdocx.CustomFileTime("2026-10-31T00:00:00Z")
    doc.save(tmp_path / "c.docx")
    back = rdocx.Document.open(tmp_path / "c.docx").custom_properties
    assert dict(back.items()) == {"Client": "Acme", "Revision": 3, "Due": "2026-10-31T00:00:00Z"}


def _toc_document(instruction, outline_paragraphs=()):
    """Chapter 1 and 2 (Heading 1), each with two Heading 2 sections, then plain paragraphs carrying a direct
    outline level, and a TOC field at the top whose instruction is `instruction`."""
    doc = rdocx.Document()
    for i in (1, 2):
        doc.add_heading(f"Chapter {i}", level=1)
        for j in (1, 2):
            doc.add_heading(f"Section {i}.{j}", level=2)
            doc.add_paragraph("Body text.")
    for text, level in outline_paragraphs:
        p = doc.add_paragraph(text)
        p.replace_xml(p.xml.replace("<w:r>", f'<w:pPr><w:outlineLvl w:val="{level - 1}"/></w:pPr><w:r>', 1))
    doc.insert_toc(0, max_level=1)
    [field] = [p for p in doc.paragraphs if 'TOC \\o "1-1" \\h' in p.xml]
    field.replace_xml(field.xml.replace('TOC \\o "1-1" \\h', instruction))
    return doc


def _toc_entries(doc):
    rep = doc.rebuild_toc()
    assert list(rep.diagnostics) == []
    return [p.text.rsplit("\t", 1)[0].strip() for p in doc.paragraphs if (p.style or "").upper().startswith("TOC")]


@pytest.mark.gap("docx-toc-outline-switch-ignores-depth")
@pytest.mark.parametrize("instruction", ['TOC \\o "1-1" \\h \\z \\u', 'TOC \\o "1-1" \\u'])
def test_toc_depth_holds_with_outline_switch(instruction):
    """`\\o "1-N"` lists heading levels 1 to N only, with `\\u` as without it (`TOC \\o "1-1" \\h` lists the two
    chapters): since 2a88a8f5 (F-283) `\\u` reads the outline level a heading style gives, and the Heading 2
    entries come in under `\\o "1-1" \\u` (Word lists Heading 1 only)."""
    assert _toc_entries(_toc_document(instruction)) == ["Chapter 1", "Chapter 2"]


@pytest.mark.gap("docx-toc-outline-level-beyond-depth")
def test_toc_outline_switch_stays_within_depth():
    """`\\u` adds the paragraphs whose own outline level lies within `\\o "1-N"`, never beyond: Word for Mac
    lists the level-1 paragraph and leaves out the level-2 one; rdocx lists both, before 2a88a8f5 as well."""
    doc = _toc_document('TOC \\o "1-1" \\h \\z \\u', [("Outline two", 2), ("Outline one", 1)])
    assert _toc_entries(doc) == ["Chapter 1", "Chapter 2", "Outline one"]
