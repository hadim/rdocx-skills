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
# html-space-doubled-at-span-edge (#359) closed with integration/open-prs-2026-10-12-3.
GAPS = {
    "font-dir-alias-below-embedded-face": "a font_dir / --font-dir file counts as the family its file name gives, "
    "but rpptx's renders still draw the face the deck embeds for that typeface; a font whose own family is the "
    "typeface wins",
}
