# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Slide header and footer fields in a large deck [pptx-edit-cost-grows-with-deck] | Of the 248 public mutating calls, three still cost a time that grows with the whole deck: a slide's `header_footer.slide_number`, `date` and `date_format` (`tests/test_pptx_scaling.py`). Every other edit costs the same at 5 slides and at 20, pictures, notes, comments, removals, slide imports, transitions, layouts, themes, charts and document properties included | Nothing to change for a few slides; on every slide of a large deck, time the first calls before setting the rest |

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
