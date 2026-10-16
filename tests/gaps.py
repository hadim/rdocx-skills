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
GAPS = {
    # upstream: tensorbee/rdocx#379
    "docx-bare-int-length-saved-as-zero": "a bare int is EMU, and font.size = 12, space_after = 6, "
    "character_spacing = 2, a margin of 1440 and line_spacing = 2 (python-docx's double spacing) are accepted and "
    "saved as about 0 instead of raising and naming Pt or Twips",
    # upstream: tensorbee/rdocx#385
    "docx-chart-read-back": "a Word document's charts cannot be read back: no doc.charts (type, categories, series)",
    # upstream: tensorbee/rdocx#385
    "docx-content-control-value-like-word": "setting a content control's value keeps w:showingPlcHdr, writes no "
    "w:fullDate for a date, writes a drop-down value as given (not its display text, not checked against the "
    "items), and an unknown tag returns 0 instead of raising KeyError",
    # upstream: tensorbee/rdocx#385
    "docx-display-equation": "add_equation inserts inline equations only: no display option (m:oMathPara)",
    # upstream: tensorbee/rdocx#385
    "docx-invalid-values-accepted": "values Word ignores or misreads are accepted silently instead of raising: "
    "doc.save(\"x.pdf\") writes a DOCX under that name, a second tab stop at one position, a page break in a header "
    "or footer, a malformed language tag (font.language = \"not a tag!\"), a font face without its name table",
    # upstream: tensorbee/rdocx#385
    "docx-list-numbering-options": "a numbered list cannot restart its count (add_numbered_list_item(restart=True), "
    "restart_numbering) and a list level takes no font (ListLevel(font=))",
    # upstream: tensorbee/rdocx#385
    "docx-note-reference-and-removal": "a footnote or endnote reference is written without Word's FootnoteReference "
    "/ EndnoteReference style (Word shows it full size), and a note cannot be removed once a later edit retired its "
    "handle (no remove_footnote(id) / remove_endnote(id))",
    # upstream: tensorbee/rdocx#385
    "docx-raw-xml-unchecked": "replace_xml accepts an unknown w: element and a foreign element that mc:Ignorable "
    "does not cover, which make Word report the file as damaged, instead of raising and leaving it unchanged",
    # upstream: tensorbee/rdocx#384
    "docx-validate-paraid-in-content-control": "rdocx validate and doc.validate() fail (\"main document paragraph "
    "owners do not match validation scope\") on a document with a paragraph carrying w14:paraId or w14:textId "
    "inside a block content control, as Word writes every table of contents",
    # upstream: tensorbee/rdocx#379
    "pptx-bare-int-spacing-saved-as-zero": "a bare int is EMU, and paragraph.space_before = 6, space_after = 6, "
    "run.font.spacing = 2 and line_spacing = 0 (or Pt(0)) are accepted and saved as 0 (spcPts val=\"0\", spc=\"0\", "
    "spcPct val=\"0\") instead of raising and naming Pt",
    # upstream: tensorbee/rdocx#385
    "pptx-chart-type-family-only": "Chart.chart_type of a chart rpptx does not write reads its plot family only: a "
    "stacked or horizontal bar reads as COLUMN_CLUSTERED (51), a line without markers as 65 (python-pptx's "
    "LINE_MARKERS), an exploded pie as PIE, instead of its python-pptx type or None",
    # upstream: tensorbee/rdocx#385
    "pptx-invalid-values-accepted": "values PowerPoint cannot honour are accepted silently instead of raising: "
    "prs.save(\"x.pdf\") writes a PPTX package under that name, slide_width or slide_height outside 1 to 56 inches is "
    "written out of the schema range, flip_h or flip_v on a table or chart frame, a table style_id that is neither "
    "built in nor in the deck",
    # upstream: tensorbee/rdocx#385
    "pptx-raw-xml-unchecked": "element.replace_xml accepts an unknown or misplaced p: or a: element (an a:p directly in "
    "p:sp, an unknown element in a:p, a:prstGeom or p:sld), which validate() does not flag either, instead of "
    "raising and leaving the deck unchanged",
    # upstream: tensorbee/rdocx#385
    "pptx-render-ignores-slide-show-master-shapes": "rpptx's renders (to_pdf, the PNG renders, convert, render) still "
    "draw the master's shapes, a logo for example, on a slide whose show_master_shapes is False; the file is right "
    "(showMasterSp=\"0\", which PowerPoint and LibreOffice honour)",
    # upstream: tensorbee/rdocx#385
    "pptx-replace-map-zero-match": "rpptx replace --map writes its output when a pair without \"expect\" finds "
    "nothing, where rpptx replace -p refuses it (\"no replacements found\")",
    # upstream: tensorbee/rdocx#379 (the paragraph._p hint is not in it)
    "python-docx-names": "python-docx names rdocx lacks: add_heading, add_page_break, insert_paragraph_before, the "
    "style= keywords of add_paragraph and add_table, Run.bold / italic / underline, doc.sections[i].header and its "
    "settable margins, add_section, Table.add_column and columns, settable TabStop, WD_BREAK, WD_COLOR_INDEX, "
    "WD_SECTION, MSO_COLOR_TYPE, styles[name], streams for open and save, dict-like custom_properties, add_picture "
    "with one dimension; paragraph._p raises a bare AttributeError instead of naming xml / replace_xml",
    # upstream: tensorbee/rdocx#379 (not in it: slide_id, ColorFormat.type and theme_color returning the enums,
    # MSO_LANGUAGE_ID, language_id, shape._element, text_frame._txBody)
    "python-pptx-names": "python-pptx names rpptx lacks: Slide.has_notes_slide and notes_slide.notes_text_frame, "
    "Slide.slide_id, MSO_COLOR_TYPE and MSO_THEME_COLOR_INDEX in rpptx.enum.dml (ColorFormat.type and theme_color read strings such "
    "as \"rgb\" and \"accent2\"), rpptx.enum.lang.MSO_LANGUAGE_ID and Font.language_id; prs.slides._sldIdLst, "
    "shape._element and text_frame._txBody raise a bare AttributeError instead of naming the rpptx way",
}
