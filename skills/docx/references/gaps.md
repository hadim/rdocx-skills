# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and section properties: `references/python-api.md`, "Raw XML").

## Headers and footers

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A picture in a header or footer [header-footer-picture] | `run.add_picture` in a header or footer run raises `NotImplementedError`, and raw XML is refused there too: a logo cannot be added from Python | write the header text with rdocx, save, then add the picture with python-docx (`section.header.paragraphs[0].add_run().add_picture(path, width=Inches(1))`) on the saved file, and reopen it with rdocx |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
