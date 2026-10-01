# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Tracked changes

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A table or row that a tracked deletion removes [accepted-view-deleted-rows] | `rdocx text`, `text --json`, Markdown, HTML and the accepted PDF keep it: an empty row line, empty cells, an empty table, or blank space in the PDF. `accept_all()` removes it correctly | accept a copy (`rdocx revision accept F -o COPY`, or `accept_all()` on a copy), then read, convert or render the copy. No fallback needed |
