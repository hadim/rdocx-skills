# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Edits of a large deck [pptx-edit-cost-grows-with-deck] | 107 of the 248 public mutating calls rewrite and reparse the whole package, so their cost grows with the deck (`tests/test_pptx_scaling.py` lists them): notes, comments, removals, `add_slide_at` and slide imports, `replace_image`, `insert_picture`, media, cell fills, `flip_h`, `align`, connections, click actions, `fit_text`, `add_field`, transitions, header and footer, layouts, theme, fonts, charts and document properties: about 150 ms a call at 5 slides of 1 MB pictures, 550 ms at 20. Constant: adding a slide, a shape, a table, a picture, a chart or a group, `duplicate` and `move`, text (`text`, `add_paragraph`, `add_run`, replacements), table and cell edits (`cell.text`, merges, rows, columns, table flags and style), `alt_title`, `name`, crops, position and size, fills, lines, paragraph and font formatting | Set the notes, comments, cell fills and properties before adding pictures, and add the pictures last: each MB in the package slows the calls that rewrite it. A deck that needs hundreds of such calls over about 20 slides: build or update it with python-pptx and check it with rpptx (`validate`, renders) |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
