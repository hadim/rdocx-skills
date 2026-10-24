# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

No gap is open on this pin: the last one the suite tracked for rpptx, edits whose cost grew with the deck, closed
with tensorbee/rdocx#413 and #416 (every public mutating call costs the same at 5 slides and at 20).

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
