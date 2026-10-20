"""pptx: fixes the pinned integration build (integration/open-prs-2026-10-20) carries and upstream `main` does not
have yet, which the skills rely on. Each test names the pull request that carries it: delete the test once upstream
main has it. Each test builds its own input."""
import pytest
import rpptx
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches


# carried by tensorbee/rdocx#393, delete once upstream main has it
def test_appends_keep_earlier_paragraph_and_run_handles():
    """add_run, add_paragraph and add_field keep earlier paragraph and run handles live; only a runs collection
    taken before an append goes stale (python-api.md, Handles)."""
    prs = rpptx.Presentation()
    frame = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_textbox(0, 0, Inches(3), Inches(1)).text_frame
    paragraph = frame.paragraphs[0]
    first = paragraph.add_run("a")
    runs = paragraph.runs
    second = paragraph.add_run("b")
    frame.add_paragraph().text = "next"
    paragraph.add_field("slidenum")
    assert (first.text, second.text, paragraph.text[:2]) == ("a", "b", "ab")
    with pytest.raises(rpptx.StaleElementError):
        len(runs)


# carried by tensorbee/rdocx#394, delete once upstream main has it
def test_placeholders_after_a_gap_clear_and_remove_style():
    """slide.placeholders has a len and iterates past a missing idx; TextFrame.clear() keeps the first paragraph
    and its formatting; Shape.remove_style() drops the p:style add_shape writes; MSO_SHAPE.UP_ARROW exists."""
    prs = rpptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.remove(slide.placeholders[0])
    assert len(slide.placeholders) == 1
    assert [ph.placeholder_format.idx for ph in slide.placeholders] == [1]
    frame = slide.shapes.add_textbox(0, 0, Inches(3), Inches(1)).text_frame
    frame.text = "one"
    frame.add_paragraph().text = "two"
    frame.paragraphs[0].font.bold = True
    frame.clear()
    assert len(frame.paragraphs) == 1 and frame.text == "" and frame.paragraphs[0].font.bold is True
    arrow = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, 0, 0, Inches(1), Inches(1))
    assert "p:style" in arrow.element.xml
    arrow.remove_style()
    assert "p:style" not in arrow.element.xml
