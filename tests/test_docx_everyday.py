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
