# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

No gap is open on this pin: every docx step the skill describes runs in rdocx. Before reaching for lxml,
try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs, tables, cells and section
properties: `references/python-api.md`, "Raw XML").

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
