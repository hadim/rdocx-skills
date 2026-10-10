# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and section properties: `references/python-api.md`, "Raw XML").

## Headers and footers

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A picture in a header or footer [header-footer-picture] | `run.add_picture` in a header or footer run raises `NotImplementedError`, and raw XML is refused there too: a logo cannot be added from Python | write the header text with rdocx, save, then add the picture with python-docx (`section.header.paragraphs[0].add_run().add_picture(path, width=Inches(1))`) on the saved file, and reopen it with rdocx |

## Reading

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A tab inside a run is dropped from story text [story-text-tab] | `StoryItem.text` in every story, the header and footer text of plain `rdocx text` and the `stories` of `rdocx text --json` join `title<tab>page` into `titlepage` | read through `Paragraph.text` (`section.footer.paragraphs[0].text` keeps `\t`) or `docx_ops.all_text(F)` / `docx_ops.py text F`, which keep it |

## Fields and the table of contents

| Gap | What happens | Workaround / fallback |
|---|---|---|
| An empty line above the first entry of a rebuilt TOC [toc-rebuild-empty-paragraph] | when the TOC field begins in its first entry and ends in its last, `rebuild_toc()` moves the begin and the end into two empty paragraphs of their own | entries, links and page numbers are right; no workaround is tested: tell the user when the layout matters |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
