# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Edits of a large deck [pptx-edit-cost-grows-with-deck] | 159 of the 248 public mutating calls cost a time that grows with the whole deck (`tests/test_pptx_scaling.py` lists them). Adding a shape, a slide, a run or a paragraph, setting text, table and cell edits, `alt_title`, `name`, crops: about 10 ms a call at 5 slides, 60 ms at 20, and faster than linearly beyond. Notes, comments, removals, `replace_image`, cell fills, `flip_h`, `align`, transitions, header and footer, layouts, theme, charts and document properties also rewrite the whole package: about 150 ms a call at 5 slides of 1 MB pictures, 550 ms at 20. Building an 80-slide deck of tables and pictures takes minutes, python-pptx seconds. Constant: `add_textbox`, position and size, `alt_text`, shape fills, lines, colours and shadows, paragraph and font formatting, `run.text`, column widths and row heights, cell margins and anchors | Over about 20 slides, build or update the deck with python-pptx and check it with rpptx (`validate`, renders). In rpptx, remove the slides you do not need first, set the notes, cell fills and properties before adding pictures, and add the pictures last: each MB in the package slows the calls that rewrite it |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
