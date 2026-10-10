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

# The nine gaps of the v0.16.0 pin (tensorbee/rdocx#295 to #301) closed with the integration build of 2026-10-10.
GAPS = {
    # docx: headers and footers
    # pptx: Python
    "font-spacing-bare-int": "font.spacing = 2 (a bare int, read as EMU) writes spc=\"0\" without a word, where the other length setters raise naming Pt",
    "header-footer-stale-shape": "slide.header_footer.slide_number = False removes the slide-number placeholder but does not retire a held handle to it when it is the last shape: it reads None and a write raises RpptxError instead of StaleElementError",
    "header-footer-picture": "a picture cannot be added to a header or footer from Python: Run.add_picture in a header or footer run raises NotImplementedError, and raw XML (Paragraph.xml / replace_xml) is refused there too",
}
