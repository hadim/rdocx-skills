# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

| Gap | What happens | Workaround | Fallback |
|---|---|---|---|
| Built-in table styles [pptx-builtin-table-styles] | a table whose style is one of PowerPoint's built-in styles referenced by GUID, not defined in the file (python-pptx writes such tables), renders unstyled | set fills on the cells you need | render with LibreOffice when the look matters |
| Line pitch [pptx-line-pitch] | a paragraph with percentage line spacing (`a:lnSpc/a:spcPct`, python-pptx `line_spacing = 1.0`) is laid out at that percentage of the font size: 100 % gives 1.0 em where LibreOffice gives 1.2 em, less than the glyphs' own height. Lines overlap in the render, and `text_layout()` reports a `height` about 17 % short: a frame it says fits can overflow in another application | for frames with percentage spacing, compare `f.height * 1.2` with `f.usable.height` | render in LibreOffice for the final check of a tight frame |
| Line breaks [pptx-line-breaks] | in a paragraph whose direction is set (`rtl="0"`, which python-pptx's template sets in its text styles), lines break at any word boundary: before a comma, before a space (the next line starts with it) or before the hyphen of a compound. `text_layout()` and the render agree with each other, not with other applications | read `text_layout()` lines: one that starts with punctuation, a space or a hyphen is a wrong break; keep a width margin (`width_factor=0.95`) | none: say so when the break matters |

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
