"""Known gaps of the pinned rdocx build: bugs and missing features, each tied to its upstream ticket.

A test marked `@pytest.mark.gap("key")` asserts the behaviour we want. It is collected as a strict xfail:
while the gap is open the test "xfails"; the day a new pin closes it, the test XPASSes, which fails the run
on purpose. Then: remove the marker (the test becomes a regression guard), delete the entry below, and
update the skills' `references/gaps.md`.

A gap about a missing API calls the name we expect it to get. When the API lands under another name, the
test keeps failing with AttributeError: adapt it to the real name, and it then reports the gap as closed.

`ticket` is the upstream issue (tensorbee/rdocx). "lot6-NN" is the draft number of the 27/09/2026 batch and
"filed 27/09" a ticket filed by hand that day, both waiting for their GitHub number; "not filed yet" is a gap
found by the acceptance review that still needs its ticket.
"""

GAPS = {
    # docx: identity attributes and producer traits (lot6-01, lot6-02)
    "toc-rsid-field-runs": ("rebuild_toc fails on a fresh open when a TOC field run carries w:rsid*", "lot6-01"),
    "compare-sdt-id": ("compare() refuses a pair whose content controls differ only by w:id or w:tag", "lot6-01"),
    "tr-identity-lost": ("w:rsidR, w:rsidTr and w14:paraId on w:tr are dropped by any modelled edit", "lot6-01"),
    "sdt-replace": ("try_replace_text and `rdocx replace` miss text inside content controls", "lot6-02"),
    "sdt-text-cli": ("`rdocx text` skips body-level content controls", "lot6-02"),
    "compare-own-save-noise": ("compare() of a file against its rdocx-edited copy reports rdocx's re-serialisation", "lot6-02"),
    "empty-comments-reserialised": ("a no-op save rewrites an empty comments.xml and compare() then refuses", "lot6-02"),
    "ignorable-undeclared": ("a save leaves an mc:Ignorable prefix undeclared in comments.xml", "lot6-02"),
    "compare-packed-fields": ("compare() refuses a file against its copy after update_page_fields on packed fields", "lot6-02"),
    # docx: comparison options (lot6-03)
    "compare-granularity": ("compare() works at whole-run granularity; word granularity is not exposed", "lot6-03"),
    "compare-comments": ("compare() refuses a pair whose comments differ", "lot6-03"),
    "compare-rebuilt-toc": ("compare() refuses a pair after rebuild_toc on the edited side", "lot6-03"),
    # docx: layout (lot6-04)
    "line-gap": ("single line height leaves out the font's line gap (Calibri 1.000 em, Word 1.221 em)", "lot6-04"),
    "picture-line-spacing": ("a line holding an inline picture is multiplied by proportional spacing", "lot6-04"),
    # docx: API (lot6-05 to lot6-08, lot6-10, lot6-13)
    "split-run-index": ("Document.split_run counts paragraphs only, so it splits the wrong one after a table", "lot6-05"),
    "save-not-atomic": ("save() writes the target in place instead of a temporary file and a rename", "lot6-06"),
    "revisions-main-story": ("Document.revisions lists the main story only", "lot6-07"),
    "cli-broken-pipe": ("the CLIs panic when their standard output is closed early", "lot6-08"),
    "cli-convert-overwrites": ("`rdocx convert --to pdf|md|html`, `rpptx convert --to pdf` and `rpptx thumbnail` overwrite any existing output, the input included", "filed 27/09"),
    "docx-paragraph-text-setter": ("no text setter on Paragraph", "lot6-10"),
    "docx-hyperlink-retarget": ("no way to change or remove an existing hyperlink", "lot6-10"),
    "docx-picture-resize": ("no way to resize an existing picture", "lot6-10"),
    "add-picture-sdt-default-ns": ("add_picture fails when the body has a content control and the root declares a default namespace", "filed 27/09"),
    "docx-python-tables": ("table merge, borders, shading, widths and row height are not bound in Python", "lot6-10"),
    "docx-python-styles": ("styles and numbering definitions cannot be created from Python", "lot6-10"),
    "docx-python-sections": ("Section is read-only in Python", "lot6-10"),
    "docx-python-bookmarks": ("bookmarks and field insertion are not bound in Python", "lot6-10"),
    "docx-replace-contract": ("no expected-count contract on Python replacements (try_replace_text returns a count only)", "lot6-10"),
    "pdf-text-ligatures": ("the PDF text layer maps glyphs by position, so ligatures garble Calibri text", "lot6-13"),
    # docx: found by the acceptance review of 27/09/2026
    "comment-runposition-sdt": ("comment run positions (Python and CLI) skip runs inside inline content controls and tracked insertions, so comments anchor on the wrong text", "filed 27/09"),
    "replace-footnotes": ("try_replace_text and `rdocx replace` skip footnotes and endnotes", "not filed yet"),
    "replace-header-footer-tables": ("try_replace_text and `rdocx replace` skip tables inside headers and footers", "not filed yet"),
    "textbox-alternate-content": ("a Word text box (mc:AlternateContent) is missing from story_items and counted twice by replacement (Choice and Fallback)", "not filed yet"),
    "text-wrapped-runs": ("text inside w:fldSimple, w:smartTag and inline w:customXml is missing from Paragraph.text, text --json and replacement", "not filed yet"),
    "replace-tracked-insertions": ("try_replace_text and `rdocx replace` do not see text inside a tracked insertion (w:ins)", "not filed yet"),
    "text-cli-tracked-insertions": ("plain `rdocx text` drops the text of tracked insertions (it shows neither the accepted nor the original view)", "not filed yet"),
    "text-other-stories": ("`rdocx text`, `text --json` and `convert --to md` leave out headers, footers, footnotes and text boxes", "not filed yet"),
    "validate-parts": ("`rdocx validate` passes a truncated header or footer part and a dangling style id", "not filed yet"),
    "style-id-unchecked": ("Paragraph.style accepts an id the package does not define (or a style name) without an error", "not filed yet"),
    "docx-core-properties": ("no API to read or write core properties (title, author, dates)", "not filed yet"),
    "comment-date-cli": ("`rdocx comment add` and `comment reply` have no --date option, so CLI comments are undated", "not filed yet"),
    "new-document-styles": ("a new Document() has only the Normal and Heading1 styles", "not filed yet"),
    "template-save-as-document": (".dotx / .potx saved as .docx / .pptx keep the template content type", "not filed yet"),
    # pptx (lot6-09, lot6-11, lot6-12)
    "pptx-run-text-stale": ("setting Run.text invalidates every handle of the presentation", "lot6-09"),
    "pptx-duplicate-slide": ("duplicate_slide exists in Rust, not bound in Python", "lot6-11"),
    "pptx-replace-python": ("replace_text / try_replace_text are not bound in Python", "lot6-11"),
    "pptx-comment-resolve-python": ("resolve_comment and remove_comment are not bound in Python", "lot6-11"),
    "pptx-inherited-geometry": ("placeholders that inherit their geometry read left/top/width/height as None", "lot6-11"),
    "pptx-group-population": ("add_group_shape returns an empty group that cannot be populated", "lot6-11"),
    "pptx-table-rows": ("no way to add or remove table rows and columns", "lot6-11"),
    "pptx-zorder": ("no way to change a shape's z-order", "lot6-11"),
    "pptx-builtin-table-styles": ("built-in table styles referenced by GUID are not applied at render", "lot6-11"),
    "pptx-pdf-background": ("the PDF export drops solid slide backgrounds that the PNG export draws", "lot6-12"),
    "pptx-gradient-optional-attrs": ("a gradient whose a:lin has no ang (python-pptx writes one) refuses the file", "lot6-12"),
    "pptx-hyperlinks": ("no hyperlink API on runs or shapes", "not filed yet"),
}
