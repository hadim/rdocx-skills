# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Formatting the text of a table cell [table-cell-text-frame] | `Cell` has `text` but no `text_frame`: no font, size, colour, bold or alignment for a cell's text | fallback: format the cells' runs with python-pptx (`cell.text_frame.paragraphs[0].runs[0].font`) on the file rpptx saved, then reopen it with rpptx; do not redraw the table as shapes |
| `font.color.rgb` [font-color-rgb-pptx] | python-pptx's `run.font.color.rgb = RGBColor(...)` raises `AttributeError` (`font.color` is None or a hex string) | `run.font.color = RGBColor(...)`; fills, lines and shadows do take `.rgb` as in python-pptx |
| Fonts from a folder [render-font-dir-pptx] | no `--font-dir` on `rpptx convert`, `render` or `thumbnail`, and no `fonts=` / `font_dir=` on `to_pdf`, `render_slide_to_png`, `render_all_slides` or `text_layout`: a font that is not installed renders, and is measured for the fit check, in a fallback | install the deck's fonts for the user first, or say that the render and the fit check use a fallback font |

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
