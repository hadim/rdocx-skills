# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

No gap is open on this pin: the last one the suite tracked for rdocx, edits whose cost grew with the body, closed
with tensorbee/rdocx#415 and #417 to #422.

Edit costs of a large document, accepted: an edit still costs a little more in a larger body. At 1000 paragraphs as
Google Docs writes them (1.7 MB of XML, measured on Linux), `insert_paragraph`, `add_paragraph` and `remove_content`
take 8 to 11 ms a call, `p.text` 27 ms, `run.text` 0.14 s, the replacements (`p.replace_text`, `cell.replace_text`,
`replace_text_at`) 0.17 to 0.18 s and `add_picture` 0.24 s (in a cell too): one replacement in each of those
paragraphs takes about 3 minutes, a `run.text` on each of their 6000 runs about 14. Calls made once or a few times
per document cost more: `move_comment` 2 s, `add_heading_hyperlink` 1.6 s, `move_content` and internal hyperlinks
about 1 s, the section and page calls (`remove_section`, `create_section_story`, `link_section_story`,
`unlink_section_story`, `is_linked_to_previous`, `set_page_color`) 0.5 to 0.7 s, accepting or rejecting a revision
by id 0.6 s, `insert_document` 0.23 s, `set_header` and `set_footer` 0.16 s. Make many replacements in one batch
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
