"""docx, reading and checking: what the pinned build gets wrong, as strict xfails (tests/gaps.py)."""
import docx
import pytest
import rdocx
from docx.oxml import parse_xml

from conftest import run

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


@pytest.mark.gap("docx-validate-paraid-in-content-control")
def test_validate_takes_a_paragraph_id_inside_a_content_control(rdocx_cli, tmp_path):
    """Word writes w14:paraId on every paragraph, those inside a block content control too (a table of contents, a
    Google Docs wrapper): the document is valid."""
    d = docx.Document()
    d.add_paragraph("Before")
    d.element.body.insert(1, parse_xml(f'<w:sdt xmlns:w="{W}" xmlns:w14="{W14}"><w:sdtPr><w:id w:val="1"/></w:sdtPr>'
                                       '<w:sdtContent><w:p w14:paraId="388B6825"><w:r><w:t>Inside</w:t></w:r></w:p>'
                                       '</w:sdtContent></w:sdt>'))
    d.save(tmp_path / "v.docx")
    res = run([rdocx_cli, "validate", tmp_path / "v.docx"])
    assert res.returncode == 0, res.stdout + res.stderr
    assert rdocx.Document.open(tmp_path / "v.docx").validate().is_valid
