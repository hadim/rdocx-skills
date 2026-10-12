# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.xml` / `replace_xml`, also on text frames,
slides and layouts: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Renders of an embedded font [render-embedded-font] | `to_pdf()`, the PNG renders and `rpptx convert` / `render` draw a font the deck embeds with a substitute (Carlito, for example); LibreOffice and PowerPoint show the embedded face | give the font file to the render as well: `font_dir=` / `--font-dir DIR` with the `.ttf` in `DIR` |
| Source formatting edited in memory [import-stale-source-formatting] | `import_slide(s)(..., keep_source_formatting=True)` brings the source's layout, master and theme as the source was opened: a layout renamed or a theme colour changed in memory before the import is lost (the slide's own content is current) | reopen the source from its bytes first: `other = rpptx.Presentation.from_bytes(other.to_bytes())`, or save it and import from the file |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
