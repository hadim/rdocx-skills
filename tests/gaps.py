"""Known gaps of the pinned rdocx build: bugs and missing features that the skills work around or leave to
another tool.

A test marked `@pytest.mark.gap("key")` asserts the behaviour we want. It is collected as a strict xfail:
while the gap is open the test "xfails"; the day a new pin closes it, the test XPASSes, which fails the run
on purpose. Then: remove the marker (the test becomes a regression guard), delete the entry below, and
update the skills' `references/gaps.md` and tables.

A gap about a missing API calls the name we expect it to get. When the API lands under another name, the
test keeps failing with AttributeError: adapt it to the real name, and it then reports the gap as closed.

Which upstream issue or pull request closes a gap is not tracked here: the suite run on each new build (the
build workflow's release notes) says which gaps closed.
"""

GAPS = {
    # docx: comparison
    "compare-rebuilt-toc": "compare() refuses a pair after rebuild_toc on the edited side",
    # docx: layout
    "line-gap": "single line height leaves out the font's line gap (Calibri 1.000 em, Word 1.221 em)",
    "picture-line-spacing": "a line holding an inline picture is multiplied by proportional spacing",
    "tab-stops": "text after a custom tab stop starts 36 pt before the stop, and a right stop is laid out as a left one",
    "toc-numbered-entries": "rebuild_toc writes number, tab, title for a numbered heading, with no stop for that tab",
    # pptx
    "pptx-builtin-table-styles": "built-in table styles referenced by GUID are not applied at render",
    "pptx-line-pitch": "a:spcPct line spacing multiplies the font size: 100 % is 1.0 em (LibreOffice 1.2 em), lines overlap",
    "pptx-line-breaks": "with rtl set on the paragraph, lines break at word boundaries: before a comma, a space or a hyphen",
}
