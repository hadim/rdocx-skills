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

GAPS = {
    # docx: comparison and rendering
    # upstream: tensorbee/rdocx#255
    "compare-final-table": "compare refuses a table added or removed with a paragraph at the very end of the body: \"comparison needs an adjacent paragraph for a final paragraph change\"",
    # upstream: tensorbee/rdocx#254
    "compare-picture-change": "compare of a picture whose image changed: no revision and the old image kept, or a refusal (\"comparison acceptance does not reproduce the edited stories\") when another paragraph changes too",
    # upstream: tensorbee/rdocx#253
    "render-tracked-view": "to_pdf, render_pages and rdocx convert render only the accepted view of tracked changes; the Rust RevisionView::Tracked is not exposed",
}
