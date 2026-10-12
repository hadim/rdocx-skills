# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.xml` / `replace_xml`, also on text frames,
slides and layouts: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A font_dir file named after the typeface [font-dir-alias-below-embedded-face] | a `font_dir=` / `--font-dir` file counts as the family its file name gives (`Aileron.ttf`), but when the deck embeds a face of that typeface the embedded face still draws; a font whose own family is the typeface wins | give the font of that family itself, or render a copy with the embedded face removed (`remove_font(typeface)`) |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
