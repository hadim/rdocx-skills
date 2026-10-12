# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and section properties: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A space doubled at a span edge [html-space-doubled-at-span-edge] | `from_html`, `insert_html`, `from_markdown` and `rdocx convert` from `.html` / `.md` keep both spaces when one ends a span and one follows it, or one precedes a span and one starts it (`a <s>b </s> c`, `and <s> that</s>`, a link `[b ](url) c`): the text reads `a b  c`, where a browser draws one space | after the import, drop the leading space of a run whose previous run ends with a space (`runs[k].text = runs[k].text[1:]`) |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
