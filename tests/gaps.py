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

# The nine gaps of the v0.16.0 pin (tensorbee/rdocx#295 to #301) closed with the integration build of 2026-10-10;
# font-spacing-bare-int and header-footer-stale-shape closed with integration/open-prs-2026-10-10-3;
# story-text-tab (#342), toc-rebuild-empty-paragraph (#341) and header-footer-picture (#343 and an integration fix
# of #322 with #328) closed with integration/open-prs-2026-10-11.
# html-space-inside-emphasis (#359), import-stale-source-formatting (#357) and render-embedded-font (#361) closed
# with integration/open-prs-2026-10-12-2.
# html-space-doubled-at-span-edge (#359) and font-dir-alias-below-embedded-face (#361) closed with
# integration/open-prs-2026-10-12-3.
GAPS = {
    # upstream: tensorbee/rdocx#371
    "compare-break-only-paragraphs": "Document.compare (and rdocx compare) fails its own acceptance check when two "
    "paragraphs holding only a line break, followed by a text paragraph, become one break-only paragraph",
    # upstream: tensorbee/rdocx#370
    "compare-empty-paragraph-properties": "Document.compare (and rdocx compare) fails its own acceptance check when "
    "an empty paragraph gains paragraph properties (spacing, alignment), in the body or a header; ignore_formatting "
    "avoids it",
    # upstream: tensorbee/rdocx#368
    "compare-hyperlink-paragraph": "Document.compare (and rdocx compare) refuses a pair in which a whole paragraph "
    "holding a hyperlink is inserted or deleted, at every granularity and with the ignore options",
    # upstream: tensorbee/rdocx#369
    "compare-paragraph-next-to-table": "Document.compare (and rdocx compare) refuses some pairs in which the "
    "paragraph right before a table is inserted or deleted (\"needs an adjacent paragraph for a final paragraph "
    "change\"), for example the first paragraph of the body, or a deletion next to an edited cell",
    # upstream: tensorbee/rdocx#372 (feature request)
    "compare-word-grouping": "Document.compare at word granularity alternates one-word deletions and insertions "
    "across a rewritten stretch instead of one deletion then one insertion, which makes a rewritten paragraph "
    "unreadable",
    # upstream: tensorbee/rdocx#373 (feature request)
    "tracked-view-colours": "the tracked revision view (to_pdf, render_pages, convert and render with "
    "--revision-view tracked) draws in black only, with no colour option; raw XML refuses header and footer "
    "paragraphs, and replace_xml refuses a picture paragraph whose asvg namespace is declared on the w:p",
}
