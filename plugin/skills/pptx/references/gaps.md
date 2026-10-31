# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Control characters [pptx-control-character-round-trip] | A part holding a character XML 1.0 forbids (U+0001 to U+001F but tab, newline and carriage return) opens, and `save()` writes the part back unchanged, so PowerPoint refuses the file. `rpptx validate` reports it, naming the part | Run `rpptx validate` on a file from an unknown producer. When it names such a character, delete it from that part's bytes (Python's `zipfile`: copy every other entry as it is), open the cleaned file, and validate the output |

Every public mutating call costs the same at 5 slides and at 20 (tensorbee/rdocx#413 and #416).

A step that rpptx blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
