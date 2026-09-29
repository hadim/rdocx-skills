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
    # docx: identity attributes and producer traits
    # docx: comparison options
    "compare-granularity": "`rdocx compare` works at whole-run granularity (Python has compare(granularity=\"word\"))",
    "compare-comments": "`rdocx compare` refuses a pair whose comments differ (Python has compare(ignore_comments=True))",
    "compare-rebuilt-toc": "compare() refuses a pair after rebuild_toc on the edited side",
    # docx: layout
    "line-gap": "single line height leaves out the font's line gap (Calibri 1.000 em, Word 1.221 em)",
    "picture-line-spacing": "a line holding an inline picture is multiplied by proportional spacing",
    "tab-stops": "text after a custom tab stop starts 36 pt before the stop, and a right stop is laid out as a left one",
    "toc-numbered-entries": "rebuild_toc writes number, tab, title for a numbered heading, with no stop for that tab",
    # docx: API
    "docx-paragraph-text-setter": "no text setter on Paragraph",
    "docx-hyperlink-retarget": "no way to change or remove an existing hyperlink",
    "docx-picture-resize": "no way to resize an existing picture",
    "docx-python-styles": "styles and numbering definitions cannot be created from Python",
    # docx: found by the acceptance review
    "replace-footnotes": "try_replace_text and `rdocx replace` skip footnotes and endnotes",
    "textbox-alternate-content": "a Word text box (mc:AlternateContent) is missing from story_items and counted twice by replacement (Choice and Fallback)",
    "text-wrapped-runs": "text inside w:fldSimple, w:smartTag and inline w:customXml is missing from Paragraph.text, text --json and replacement",
    "replace-tracked-insertions": "try_replace_text and `rdocx replace` do not see text inside a tracked insertion (w:ins)",
    "text-cli-tracked-insertions": "plain `rdocx text` drops the text of tracked insertions (it shows neither the accepted nor the original view)",
    "text-other-stories": "`rdocx text`, `text --json` and `convert --to md` leave out headers, footers, footnotes and text boxes",
    "validate-parts": "`rdocx validate` passes a truncated header or footer part and a dangling style id",
    "style-id-unchecked": "Paragraph.style accepts an id the package does not define (or a style name) without an error",
    "docx-core-properties": "no API to read or write core properties (title, author, dates)",
    "comment-date-cli": "`rdocx comment add` and `comment reply` have no --date option, so CLI comments are undated",
    "new-document-styles": "a new Document() has only the Normal and Heading1 styles",
    "template-save-as-document": ".dotx / .potx saved as .docx / .pptx keep the template content type",
    # pptx
    "pptx-group-population": "add_group_shape returns an empty group that cannot be populated",
    "pptx-table-rows": "no way to add or remove table rows and columns",
    "pptx-builtin-table-styles": "built-in table styles referenced by GUID are not applied at render",
    "pptx-line-pitch": "a:spcPct line spacing multiplies the font size: 100 % is 1.0 em (LibreOffice 1.2 em), lines overlap",
    "pptx-line-breaks": "with rtl set on the paragraph, lines break at word boundaries: before a comma, a space or a hyphen",
    "pptx-duplicate-ppr": "a paragraph with two a:pPr makes the whole file refuse to open",
}
