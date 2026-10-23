"""docx: what the pinned build lacks of the everyday surface the skills cite, each as a strict xfail
(tests/gaps.py). Each test builds its own input."""
import pytest
import rdocx


@pytest.mark.gap("docx-default-line-spacing-multiple")
def test_default_line_spacing_takes_a_multiple():
    """set_default_paragraph_format takes a length (an exact pitch) but no multiple, the rule of Word's own default
    (1.08 lines); reading a multiple works."""
    doc = rdocx.Document()
    doc.set_default_paragraph_format(line_spacing=1.15)
    assert doc.default_paragraph_format.line_spacing == pytest.approx(1.15)
    assert 'w:line="276" w:lineRule="auto"' in doc.part_xml("word/styles.xml")
