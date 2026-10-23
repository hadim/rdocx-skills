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
GAPS = {
    # upstream: tensorbee/rdocx#402
    "docx-default-line-spacing-multiple": "set_default_paragraph_format(line_spacing=) takes a length (an exact "
    "pitch) and raises TypeError for a float multiple, the rule of Word's own default; set the multiple on the "
    "Normal style",
    # upstream: tensorbee/rdocx#404
    "docx-insert-document-reused-picture-swapped": "insert_document reuses the target's picture parts that hold the "
    "inserted pictures' bytes, paired in the wrong order: an inserted picture shows another one; pass "
    "policy=FragmentConflictPolicy.reuse_equivalent().with_related_part_reuse(False)",
    # upstream: tensorbee/rdocx#405
    "docx-comments-range-across-cells": "Document.comments raises RdocxError when one comment's range runs from a "
    "table cell into the next one, as Google Docs writes a comment across cells; remove_comment works, part_xml reads "
    "word/comments.xml",
    # upstream: tensorbee/rdocx#406
    "docx-part-xml-without-relationships": "part_names() and part_xml() leave out the .rels parts and "
    "[Content_Types].xml; unzip to_bytes() to read them",
    # test_docx_scaling.py: the calls of its GROWING set; an XPASS removes that call from GROWING, and this entry goes
    # when GROWING is empty
    # upstream: tensorbee/rdocx#407
    "docx-edit-cost-grows-with-body": "the call costs a time that grows with the body's XML, whatever the element "
    "read or edited; edit the structure in few calls (remove_content_range, insert_document of a document built "
    "aside)",
    # upstream: tensorbee/rdocx#408
    "docx-insert-document-duplicate-style-ids": "insert_document adds <id>Merge1 copies of the styles that the "
    "styles part defines twice, though both documents carry the same part; keep the first definition of each id in "
    "the inserted document",
    # upstream: tensorbee/rdocx#410 (the pull request of #408)
    "docx-insert-document-new-document-refused": "insert_document and import_fragment of a new Document() raise "
    "RdocxError (more than one default) when the target's Normal style differs from rdocx's own, as in documents "
    "from Word, Google Docs or python-docx; build the content in a copy of the target emptied by remove_content_range",
    # test_pptx_scaling.py: 159 of the 248 public mutating calls of rpptx 0.14.3, its GROWING set; an XPASS removes
    # that call from GROWING, and this entry goes when GROWING is empty
    # upstream: tensorbee/rdocx#403
    "pptx-edit-cost-grows-with-deck": "the call costs a time that grows with the rest of the deck: it copies the "
    "deck and rebuilds the inventory of every shape, and many calls also rewrite and reparse the whole package "
    "(about 22 ms a MB); build or update a large deck with python-pptx, add pictures last",
}
