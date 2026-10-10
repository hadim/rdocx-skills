# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.xml` / `replace_xml`, also on text frames,
slides and layouts: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A bare int for character spacing [font-spacing-bare-int] | `font.spacing = 2` is read as 2 EMU and written as `spc="0"`, without an error | always give a length: `font.spacing = Pt(2)` |
| A held footer placeholder [header-footer-stale-shape] | after `slide.header_footer.slide_number = False` (or `set_header_footer(slide_number=False)`), a handle to the removed placeholder that was the slide's last shape reads None instead of raising | re-fetch `prs.slides[i].shapes` after switching a footer element off |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
