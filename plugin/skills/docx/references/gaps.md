# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Listing custom properties [docx-custom-properties-not-listed] | `doc.custom_properties` takes `[name] = value`, `del` and `in`, and `get(name).value` reads one property, but nothing lists them: no `keys()`, `items()` or iteration | `rdocx meta get --json F`: its `custom` array gives each property's `name`, `type` and `value` |
| TOC depth with `\u` [docx-toc-outline-switch-ignores-depth] | `rebuild_toc()`, `rdocx toc rebuild` and `docx_ops.py toc` ignore the depth of `\o "1-N"` when the instruction also carries `\u` (Word's own `TOC \o "1-3" \h \z \u` does): every heading level gets an entry, with no diagnostic. A regression of this pin | take `\u` out for the rebuild and put it back (`recipes.md`, "Table of contents and page fields"); check `rep.entry_count` against the headings of levels 1 to N |
| Outline levels beyond the TOC depth [docx-toc-outline-level-beyond-depth] | under `\o "1-N" \u`, a paragraph whose own outline level is beyond N (no heading style) still gets an entry; Word leaves it out | the same recipe: without `\u` such paragraphs get no entry at all, which is right beyond N; a paragraph within 1 to N then needs a heading style, or an update in Word |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
