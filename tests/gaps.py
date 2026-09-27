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
    "toc-rsid-field-runs": "rebuild_toc fails on a fresh open when a TOC field run carries w:rsid*",
    "compare-sdt-id": "compare() refuses a pair whose content controls differ only by w:id or w:tag",
    "tr-identity-lost": "w:rsidR, w:rsidTr and w14:paraId on w:tr are dropped by any modelled edit",
    "sdt-replace": "try_replace_text and `rdocx replace` miss text inside content controls",
    "sdt-text-cli": "`rdocx text` skips body-level content controls",
    "compare-own-save-noise": "compare() of a file against its rdocx-edited copy reports rdocx's re-serialisation",
    "empty-comments-reserialised": "a no-op save rewrites an empty comments.xml and compare() then refuses",
    "ignorable-undeclared": "a save leaves an mc:Ignorable prefix undeclared in comments.xml",
    "ignorable-dropped": "an edit that rewrites document.xml, a header or a footer drops the root's mc:Ignorable",
    "compare-packed-fields": "compare() refuses a file against its copy after update_page_fields on packed fields",
    # docx: comparison options
    "compare-granularity": "compare() works at whole-run granularity; word granularity is not exposed",
    "compare-comments": "compare() refuses a pair whose comments differ",
    "compare-rebuilt-toc": "compare() refuses a pair after rebuild_toc on the edited side",
    # docx: layout
    "line-gap": "single line height leaves out the font's line gap (Calibri 1.000 em, Word 1.221 em)",
    "picture-line-spacing": "a line holding an inline picture is multiplied by proportional spacing",
    "tab-stops": "text after a custom tab stop starts 36 pt before the stop, and a right stop is laid out as a left one",
    "toc-numbered-entries": "rebuild_toc writes number, tab, title for a numbered heading, with no stop for that tab",
    # docx: API
    "split-run-index": "Document.split_run counts paragraphs only, so it splits the wrong one after a table",
    "save-not-atomic": "save() writes the target in place instead of a temporary file and a rename",
    "revisions-main-story": "Document.revisions lists the main story only",
    "cli-broken-pipe": "the CLIs panic when their standard output is closed early",
    "cli-convert-overwrites": "`rdocx convert --to pdf|md|html`, `rpptx convert --to pdf` and `rpptx thumbnail` overwrite any existing output, the input included",
    "docx-paragraph-text-setter": "no text setter on Paragraph",
    "docx-hyperlink-retarget": "no way to change or remove an existing hyperlink",
    "docx-picture-resize": "no way to resize an existing picture",
    "add-picture-sdt-default-ns": "add_picture fails when the body has a content control and the root declares a default namespace",
    "docx-python-tables": "table merge, borders, shading, widths and row height are not bound in Python",
    "docx-python-styles": "styles and numbering definitions cannot be created from Python",
    "docx-python-sections": "Section is read-only in Python",
    "docx-python-bookmarks": "bookmarks and field insertion are not bound in Python",
    "docx-replace-contract": "no expected-count contract on Python replacements (try_replace_text returns a count only)",
    "pdf-text-ligatures": "the PDF text layer maps glyphs by position, so ligatures garble Calibri text",
    # docx: found by the acceptance review
    "comment-runposition-sdt": "comment run positions (Python and CLI) skip runs inside inline content controls and tracked insertions, so comments anchor on the wrong text",
    "replace-footnotes": "try_replace_text and `rdocx replace` skip footnotes and endnotes",
    "replace-header-footer-tables": "try_replace_text and `rdocx replace` skip tables inside headers and footers",
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
    "pptx-run-text-stale": "setting Run.text invalidates every handle of the presentation",
    "pptx-duplicate-slide": "duplicate_slide exists in Rust, not bound in Python",
    "pptx-replace-python": "replace_text / try_replace_text are not bound in Python",
    "pptx-comment-resolve-python": "resolve_comment and remove_comment are not bound in Python",
    "pptx-inherited-geometry": "placeholders that inherit their geometry read left/top/width/height as None",
    "pptx-group-population": "add_group_shape returns an empty group that cannot be populated",
    "pptx-table-rows": "no way to add or remove table rows and columns",
    "pptx-zorder": "no way to change a shape's z-order",
    "pptx-builtin-table-styles": "built-in table styles referenced by GUID are not applied at render",
    "pptx-pdf-background": "the PDF export drops solid slide backgrounds that the PNG export draws",
    "pptx-gradient-optional-attrs": "a gradient whose a:lin has no ang (python-pptx writes one) refuses the file",
    "pptx-hyperlinks": "no hyperlink API on runs or shapes",
    "pptx-line-pitch": "a:spcPct line spacing multiplies the font size: 100 % is 1.0 em (LibreOffice 1.2 em), lines overlap",
    "pptx-line-breaks": "with rtl set on the paragraph, lines break at word boundaries: before a comma, a space or a hyphen",
    "pptx-duplicate-ppr": "a paragraph with two a:pPr makes the whole file refuse to open",
}
