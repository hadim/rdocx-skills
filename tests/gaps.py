"""Known gaps of the pinned rdocx build: bugs and missing features that the skills work around or leave to
another tool.

A test marked `@pytest.mark.gap("key")` asserts the behaviour we want. It is collected as a strict xfail:
while the gap is open the test "xfails"; the day a new pin closes it, the test XPASSes, which fails the run
on purpose. Then: remove the marker (the test becomes a regression guard), delete the entry below, and
update the skills' `references/gaps.md` and tables.

A gap about a missing API calls the name we expect it to get. When the API lands under another name, the
test keeps failing with AttributeError: adapt it to the real name, and it then reports the gap as closed.

When a gap is reported upstream, a comment on its entry names the issue (`# upstream: owner/repo#N`). The
issue's state decides nothing: the suite run on each new pin (ci.yml's job summaries) says which
gaps closed.
"""

# The nine gaps of the v0.16.0 pin (tensorbee/rdocx#295 to #301) closed with the integration build of 2026-10-10;
# font-spacing-bare-int and header-footer-stale-shape closed with integration/open-prs-2026-10-10-3;
# story-text-tab (#342), toc-rebuild-empty-paragraph (#341) and header-footer-picture (#343 and an integration fix
# of #322 with #328) closed with integration/open-prs-2026-10-11.
# html-space-inside-emphasis (#359), import-stale-source-formatting (#357) and render-embedded-font (#361) closed
# with integration/open-prs-2026-10-12-2.
# html-space-doubled-at-span-edge (#359) and font-dir-alias-below-embedded-face (#361) closed with
# integration/open-prs-2026-10-12-3.
# compare-hyperlink-paragraph (#368), compare-paragraph-next-to-table (#369), compare-empty-paragraph-properties
# (#370), compare-break-only-paragraphs (#371), compare-word-grouping (#372, through #374) and tracked-view-colours
# (#373, through #375) closed with integration/open-prs-2026-10-14.
# 2026-10-16: the pin moved from integration/open-prs-2026-10-14 to tensorbee/rdocx main 3a3b0de8, which rewrote
# every carried pull request: the behaviour tests that pass there were deleted, and the gaps below are what main
# still lacks or got wrong.
# Main d44a69af (S94, tensorbee/rdocx#379) closed pptx-bare-int-spacing-saved-as-zero and the python-pptx notes names,
# and most python-docx names: add_heading, add_page_break, insert_paragraph_before, style=, Run.bold / italic /
# underline, WD_BREAK, WD_COLOR_INDEX, WD_SECTION, streams; Section.header, add_section, Table.add_column / columns,
# styles[name], add_picture with one dimension and prs.slides._sldIdLst now raise an error that names the rdocx or rpptx way, which is the contract
# #379 asked for, so they left the gaps.
# 2026-10-18: main 7496a81c (v0.16.2 and S95.11, tensorbee/rdocx#384, #385 and #386) closed every other gap of
# d44a69af: docx-bare-int-length-saved-as-zero, docx-chart-read-back, docx-content-control-value-like-word,
# docx-display-equation, docx-invalid-values-accepted, docx-list-numbering-options, docx-note-reference-and-removal,
# docx-raw-xml-unchecked, docx-validate-paraid-in-content-control, pptx-chart-type-family-only,
# pptx-invalid-values-accepted, pptx-raw-xml-unchecked, pptx-render-ignores-slide-show-master-shapes,
# pptx-replace-map-zero-match, python-docx-names and python-pptx-names. Of the names, style="List Bullet" on a new
# document, a settable TabStop, paragraph._p, shape._element and text_frame._txBody now raise an error that names the
# rdocx or rpptx way (KeyError or NotImplementedError), the contract #379 asked for; custom_properties became a
# mapping for writing only (below).
GAPS = {
    # not reported upstream yet
    "docx-custom-properties-not-listed": "custom_properties takes item assignment, deletion and `in`, and "
    "get(name).value reads one property, but nothing lists the properties a document holds (no keys(), items() or "
    "iteration); rdocx meta get --json lists them",
}
