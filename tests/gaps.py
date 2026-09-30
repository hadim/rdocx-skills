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
    # docx: styles
    "styles-duplicate-ids": "add_style raises \"invalid style graph: duplicate style ID\" when styles.xml repeats a style id",
}
