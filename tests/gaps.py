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
# font-spacing-bare-int and header-footer-stale-shape closed with integration/open-prs-2026-10-10-3.
GAPS = {
    # docx: headers and footers
    "header-footer-picture": "a picture cannot be added to a header or footer from Python: Run.add_picture in a header or footer run raises NotImplementedError, and raw XML (Paragraph.xml / replace_xml) is refused there too",
    # docx: reading
    # upstream: tensorbee/rdocx#339
    "story-text-tab": "a w:tab in a run is dropped from StoryItem.text in every story, from the header and footer text of plain `rdocx text` and from the stories of `rdocx text --json` (Paragraph.text keeps it as a tab)",
    # docx: fields and the table of contents
    # upstream: tensorbee/rdocx#340
    "toc-rebuild-empty-paragraph": "rebuild_toc moves the TOC field's begin out of the first entry and its end out of the last into two empty paragraphs of their own: an empty line above the first entry",
}
