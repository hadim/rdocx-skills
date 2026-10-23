# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A multiple as the default line spacing [docx-default-line-spacing-multiple] | `doc.set_default_paragraph_format(line_spacing=1.15)` raises `TypeError`: it takes a length, which writes an exact pitch (`w:lineRule="exact"`), and no multiple (`auto`), the rule of Word's own default (1.08 lines). Reading one works | `doc.set_style("Normal", line_spacing=1.15)`: the paragraphs that follow Normal, and the styles based on it, take it |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
