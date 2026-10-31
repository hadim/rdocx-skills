# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Control characters [docx-control-character-round-trip] | A part holding a character XML 1.0 forbids (U+0001 to U+001F but tab, newline and carriage return), as some exporters write, opens, and `save()` writes the part back unchanged, so Word refuses the file. `rdocx validate` reports it, naming the part, line and column | Run `rdocx validate` on a file from an unknown producer. When it names such a character, delete it from that part's bytes (Python's `zipfile`: copy every other entry as it is), open the cleaned file, and validate the output |
| Fonts rdocx does not bundle [docx-font-fallback-by-entry-point] | For a font the document names that rdocx does not bundle and the system has (DejaVu Sans on Linux), `rdocx convert` and `rdocx render` draw a bundled substitute without a message, `to_pdf()` the system font, and `page_count()` / `layout()` one or the other depending on `font_dir`: a page count taken in Python can differ from the PDF written next to it | Put the font files the document names in a folder and pass it to every call, `--font-dir DIR` and `font_dir=DIR`: then all of them draw that font and count the same pages. Check which font a PDF used with `pdffonts` |

Edit costs of a large document, accepted: an edit still costs a little more in a larger body. At 1000 paragraphs as
Google Docs writes them (1.7 MB of XML, measured on Linux), `insert_paragraph`, `add_paragraph` and `remove_content`
take 8 to 11 ms a call, `p.text` 27 ms, `run.text` 0.14 s, the replacements (`p.replace_text`, `cell.replace_text`,
`replace_text_at`) 0.17 to 0.18 s and `add_picture` 0.24 s (in a cell too): one replacement in each of those
paragraphs takes about 3 minutes, a `run.text` on each of their 6000 runs about 14. Calls made once or a few times
per document cost more: `move_comment` 2 s, `add_heading_hyperlink` 1.6 s, `move_content` and internal hyperlinks
about 1 s, the section and page calls (`remove_section`, `create_section_story`, `link_section_story`,
`unlink_section_story`, `is_linked_to_previous`, `set_page_color`) 0.5 to 0.7 s, accepting or rejecting a revision
by id 0.6 s, `insert_document` 0.23 s, `set_header` and `set_footer` 0.16 s. Edits also cost more with a larger
`styles.xml`, the body being the same: `run.text`, `try_replace_text`, `table.add_row()` and `clone_row` take 0.5 to 5
ms a call with rdocx's own 6 KB styles part, 30 to 60 ms with 2000 unused styles (419 KB; python-docx's default
template has 349 KB), and a setter reached through a table cell grows with the cell's place in the table
(`references/python-api.md`, Handles: one `replace_xml` per table). Make many replacements in one batch
(`docx_ops.replace_batch(src, out, [(old, new, n), ...])` or `rdocx replace F --map pairs.json -o OUT`). Make many
other edits in a document built aside, which is small, then insert it with one `insert_document`: a new
`rdocx.Document()`, or a copy of the target emptied by one `remove_content_range(start, end)`
(`rdocx.Document.from_bytes(doc.to_bytes())`) to start from its styles. Where the runs' formatting does not matter,
one `p.text` per paragraph costs less than a `run.text` on each run.

The form operations (`validate_form`, `fill_form`, `reset_form`, `protect_form`) check the whole document on each
call, at a cost that grows faster than the body: up to 0.1 s at 25 paragraphs, 3 to 9 s at 1000. Pass every rule and
assignment in one call.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
