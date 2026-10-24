"""Known gaps of the pinned rdocx build: bugs and missing features that the skills work around or leave to
another tool.

A test marked `@pytest.mark.gap("key")` asserts the behaviour we want. It is collected as a strict xfail:
while the gap is open the test "xfails"; the day a new pin closes it, the test XPASSes, which fails the run
on purpose. Then: remove the marker (the test becomes a regression guard), delete the entry below once no test
carries its key, and update the skills' `references/gaps.md` and tables.

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
# 2026-10-20: integration/open-prs-2026-10-20 (main b072b985 plus tensorbee/rdocx#393, #394, #395, #396, #397 and
# #399: cell runs after cell edits, rpptx handles across appends, placeholders and python-pptx behaviours, part XML,
# default paragraph format, text-box paragraph handles, range removal, linear handle walks) closed no gap; the fixes
# the skills rely on are in test_docx_carried.py and test_pptx_carried.py until upstream main has them.
# 2026-10-21: the same pin regressed on rebuild_toc: tensorbee/rdocx 2a88a8f5 (F-283) made `\\u` read the outline
# level a heading style gives, and the outline branch never honoured the `\\o` range, so `\\o "1-1" \\u` lists Heading 2
# (docx-toc-outline-switch-ignores-depth; 7496a81c passes it). docx-toc-outline-level-beyond-depth is the same missing
# range check on a paragraph's own outline level, already there in 7496a81c. Both reported as tensorbee/rdocx#400 and
# closed with integration/open-prs-2026-10-21 (the fix tensorbee/rdocx#401 on top of -10-20); the tests moved to
# test_docx_carried.py until upstream main has it.
# 2026-10-23: the pin moved to the upstream release v0.16.3 / rpptx-v0.14.3 (tensorbee/rdocx a57be365, S95.3), which
# adapted every carried pull request (#393 to #397, #399, #401): test_docx_carried.py and test_pptx_carried.py were
# deleted, their tests passing there, or for remove_content_range (start, end; RdocxError out of range) differing only
# in names. docx-custom-properties-not-listed (#387) closed: custom_properties lists (items(), iteration, len,
# [name] gives the CustomProperty). New gap docx-default-line-spacing-multiple: the carried #395 took a float multiple
# in set_default_paragraph_format, upstream's version takes lengths only.
# 2026-10-23: five findings of real use on the same pin, reported as tensorbee/rdocx#404 to #408, became the gaps
# docx-insert-document-reused-picture-swapped, docx-comments-range-across-cells, docx-part-xml-without-relationships,
# docx-edit-cost-grows-with-body (test_docx_scaling.py, the docx counterpart of test_pptx_scaling.py) and
# docx-insert-document-duplicate-style-ids. Writing the scaling test met docx-insert-document-new-document-refused,
# fixed in the pull request of #408 without an issue of its own.
# 2026-10-23: integration/open-prs-2026-10-23 (v0.16.3 plus tensorbee/rdocx#409 to #413) closed
# docx-part-xml-without-relationships (#409), docx-insert-document-duplicate-style-ids and
# docx-insert-document-new-document-refused (#410), docx-insert-document-reused-picture-swapped (#411) and
# docx-comments-range-across-cells (#412); their tests moved to test_docx_carried.py until upstream main has them.
# integration/open-prs-2026-10-23-2 (plus tensorbee/rdocx#414 and #415) closed docx-default-line-spacing-multiple
# (#414: set_default_paragraph_format, add_style and set_style take a number of lines).
# integration/open-prs-2026-10-23-5 (-2 with #415 at its final head, plus #416 and the docx edit cost series #417 to
# #420) made 69 more calls of test_docx_scaling.py flat on the four CI platforms (75 of the 252 still grow), and the
# last nine of test_pptx_scaling.py, slide transitions and header and footer setters: pptx-edit-cost-grows-with-deck
# closed (#413, #416).
# integration/open-prs-2026-10-23-6 (-5 plus #421, the fifth pull request of the series) made 61 more calls flat (14
# of the 252 still grow). 35 of them sit near FIXED, so that a noisy runner makes them XPASS on some platforms and
# runs only: they left GROWING all the same, being far under RATIO (at most 1.62 times on Linux, --runxfail).
GAPS = {
    # test_docx_scaling.py: the calls of its GROWING set; an XPASS removes that call from GROWING, and this entry goes
    # when GROWING is empty
    # upstream: tensorbee/rdocx#407
    "docx-edit-cost-grows-with-body": "the call costs a time that grows with the body's XML (replacements, pictures, "
    "moved content, headers and footers, revisions by id, comments moved, story text and hyperlinks); make "
    "many such edits in a document built aside, then one insert_document",
}
