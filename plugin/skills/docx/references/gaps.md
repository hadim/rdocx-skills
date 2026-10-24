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
take about 2 ms a call, `p.text` 17 ms, `run.text` 70 ms, the replacements (`p.replace_text`, `cell.replace_text`,
`replace_text_at`) 0.11 to 0.14 s and `add_picture` 0.19 s (0.16 s in a cell): one replacement in each of those
paragraphs takes about 2.5 minutes, a `run.text` on each of their 6000 runs about 7. Calls made once or a few
times per document cost more: `move_comment` 1.4 s, `move_content` and internal hyperlinks 0.65 to 0.75 s,
accepting or rejecting a revision by id 0.55 s, `add_heading_hyperlink` 0.44 s, `set_header` and `set_footer`
0.37 s, `insert_document` 0.2 s. Make many replacements in one batch (`docx_ops.replace_batch(src, out, [(old, new,
n), ...])` or `rdocx replace F --map pairs.json -o OUT`). Make many other edits in a document built aside, which is
small, then insert it with one `insert_document`: a new `rdocx.Document()`, or a copy of the target emptied by one
`remove_content_range(start, end)` (`rdocx.Document.from_bytes(doc.to_bytes())`) to start from its styles. Where
the runs' formatting does not matter, one `p.text` per paragraph costs less than a `run.text` on each run.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
