# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| python-pptx names [python-pptx-names] | `slide.has_notes_slide`, `slide.notes_slide.notes_text_frame`, `slide.slide_id`, `MSO_COLOR_TYPE` and `MSO_THEME_COLOR_INDEX` (`rpptx.enum.dml`), `rpptx.enum.lang.MSO_LANGUAGE_ID` and `font.language_id` do not exist; `color.type` and `color.theme_color` read strings; `prs.slides._sldIdLst`, `shape._element` and `text_frame._txBody` raise a bare `AttributeError` | `slide.notes_text` (get and set, None without notes); compare slides with `==`, ids from `rpptx text --json`; compare `color.type == "rgb"` or `"theme"` and set `theme_color = MSO_THEME_COLOR.ACCENT_2` (it reads `"accent2"`); `font.language = "fr-FR"`; `prs.slides.remove` / `move`, `element.xml` / `element.replace_xml` |
| Spacings saved as zero [pptx-bare-int-spacing-saved-as-zero] | a bare int is EMU: `paragraph.space_before = 6`, `space_after = 6`, `run.font.spacing = 2` and `line_spacing = 0` (or `Pt(0)`) are accepted and saved as 0, with no error | always write `Pt(6)`, `Pt(2)`, or a float such as 1.5 (a multiple of the line); for no extra spacing, `None` |
| Values PowerPoint cannot honour [pptx-invalid-values-accepted] | accepted without an error: `prs.save("x.pdf")` (writes a PPTX package under that name), `slide_width` or `slide_height` outside 1 to 56 inches (written out of the schema range), `flip_h` / `flip_v` on a table or chart frame, a table `style_id` that is neither built in nor in the deck | a PDF is `open("x.pdf", "wb").write(prs.to_pdf())` or `rpptx convert --to pdf`; keep the slide size from 1 to 56 inches; do not flip tables or charts; take a `style_id` from PowerPoint's built-in styles or from a table of the deck |
| Unchecked raw XML [pptx-raw-xml-unchecked] | `element.replace_xml` refuses malformed XML, a DOCTYPE, another root and a missing `r:id`, but accepts an unknown or misplaced `p:` or `a:` element (an `a:p` directly in `p:sp`, an unknown element in `a:p`), which PowerPoint cannot open, and `validate()` does not flag it | copy the structure from XML PowerPoint wrote (another deck's `element.xml`), keep each element where the schema puts it, then open the result in LibreOffice (`soffice --headless --convert-to pdf`) before delivering it |
| Chart type of a chart rpptx does not write [pptx-chart-type-family-only] | `chart.chart_type` reads the plot family only: a stacked or horizontal bar reads as `COLUMN_CLUSTERED` (51), a line without markers as 65 (`rpptx.XL_CHART_TYPE.LINE`, python-pptx's `LINE_MARKERS`), an exploded pie as `PIE` | read the type with python-pptx (`chart.chart_type` there gives the variant); rpptx's `categories`, `series` and `replace_data` are right |
| Replacement map pair without a count [pptx-replace-map-zero-match] | `rpptx replace --map` writes its output when a pair without `"expect"` finds nothing (`-p` refuses it: "no replacements found") | give every pair its `"expect"`, or check the `count` of each pair in the `--json` record |
| Renders of a slide that hides the master shapes [pptx-render-ignores-slide-show-master-shapes] | `to_pdf`, the PNG renders, `rpptx convert` and `render` still draw the master's shapes (a logo) on a slide whose `show_master_shapes` is False; the file is right (`showMasterSp="0"`): PowerPoint and LibreOffice hide them | check such a slide with LibreOffice (`soffice --headless --convert-to pdf`) |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
