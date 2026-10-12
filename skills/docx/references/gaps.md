# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and section properties: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A space before emphasis [html-space-inside-emphasis] | `from_html`, `from_markdown` and `rdocx convert` from `.html` / `.md` put the space before an underlined, struck, bold or italic span inside its run (`" this"`): Word underlines or strikes that space | after the import, move the space out: set the run's text without its leading space and add it to the end of the run before (`runs[k - 1].text += " "`) |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
