"""Known gaps of the pinned rdocx build: bugs and missing features that the skills work around or leave to
another tool.

A test marked `@pytest.mark.gap("key")` asserts the behaviour we want. It is collected as a strict xfail:
while the gap is open the test "xfails"; the day a new pin closes it, the test XPASSes, which fails the run
on purpose. Then: remove the marker (the test becomes a regression guard), delete the entry below, and
update the skills' `references/gaps.md` and tables.

A gap about a missing API calls the name we expect it to get. When the API lands under another name, the
test keeps failing with AttributeError: adapt it to the real name, and it then reports the gap as closed.

When a gap is reported upstream, a comment on its entry names the issue (`# upstream: owner/repo#N`). The
issue's state decides nothing: the suite run on each new build (the build workflow's release notes) says which
gaps closed.
"""

GAPS = {
    # docx and pptx: rendering
    # upstream: tensorbee/rdocx#296
    "render-font-dir-docx": "caller fonts reach only the PDF: rdocx convert --to png/jpeg/tiff ignores --font-dir silently, rdocx render has no --font-dir, Document.render_page_to_png / render_all_pages / render_pages take no fonts",
    # upstream: tensorbee/rdocx#296
    "render-font-dir-pptx": "rpptx has no way to give fonts: no --font-dir on convert, render or thumbnail, no fonts= on Presentation.to_pdf, render_slide_to_png, render_all_slides or text_layout",
    # upstream: tensorbee/rdocx#297
    "variable-font-bold": "a variable font given through --font-dir / font_dir renders a bold run with the regular glyphs under a -Bold font name: the wght axis is not instanced and no synthetic bold is applied",
    # docx: validation
    # upstream: tensorbee/rdocx#298
    "validate-picture-paragraph": "rdocx validate counts a paragraph that holds only a picture as an empty paragraph",
    # docx and pptx: Python colours
    # upstream: tensorbee/rdocx#299
    "colour-forms-docx": "colour arguments take an RGBColor (set_style, add_style, font.color) or a hex string (set_borders, set_border, shading) but not both, and font.color.rgb (python-docx) raises AttributeError",
    # upstream: tensorbee/rdocx#299
    "font-color-rgb-pptx": "run.font.color.rgb = RGBColor(...) (python-pptx) raises AttributeError: rpptx takes font.color = RGBColor(...) and reads back a hex string",
    # docx: tables
    # upstream: tensorbee/rdocx#300
    "table-indent": "the Python Table has no indent (w:tblInd); the Rust set_indent_checked refuses a negative indent, which Word writes and honours",
    # pptx: tables
    # upstream: tensorbee/rdocx#295
    "table-cell-text-frame": "a table Cell has text but no text_frame in Python, so the text of a cell cannot be formatted (font, size, colour, alignment)",
    # docx: editing
    # upstream: tensorbee/rdocx#301
    "replace-text-at-text-box-copy": "replace_text_at on a text box paragraph edits the DrawingML copy (mc:Choice) only; the VML fallback (mc:Fallback) keeps the old text, so the two copies of the box disagree",
}
