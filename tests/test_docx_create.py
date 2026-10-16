"""docx, creating a document: what the pinned build lacks, as strict xfails (tests/gaps.py)."""
import io

import pytest
import rdocx
from PIL import Image


@pytest.mark.gap("python-docx-names")
def test_add_picture_with_one_dimension_keeps_the_aspect_ratio():
    """As python-docx: a width alone (or a height alone) scales the other side."""
    d = rdocx.Document()
    png = io.BytesIO()
    Image.new("RGB", (20, 10)).save(png, "PNG")
    d.add_picture(png.getvalue(), "x.png", width=rdocx.Inches(2))
    d.add_picture(png.getvalue(), height=rdocx.Inches(2))
    assert [(p.width, p.height) for p in d.pictures()] == [(rdocx.Inches(2), rdocx.Inches(1)), (rdocx.Inches(4), rdocx.Inches(2))]
