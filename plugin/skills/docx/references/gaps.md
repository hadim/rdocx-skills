# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Redline by word [docx-compare-word-splits-characters] | At `granularity="word"` (`--granularity word`), each character of a changed paragraph becomes a run of its own and each deleted or inserted letter its own `w:delText` / `w:t`: the redline reads the same in Word, but its XML is about twice as large, a compare of a long document takes about ten times as long (200 changed paragraphs: 6.9 s against 0.65 s in rdocx 0.16.4), and `p.runs` of a changed paragraph holds one run per character | Keep `--granularity word` for a redline a person reads, and edit the redline's changed paragraphs by text (`replace_text`), not by run index. For a long document where time matters, compare with the default `run` granularity |
| Moving a thread [docx-move-thread-leaves-replies] | `move_comment_to_text`, `move_comment` and `rdocx comment move` move the root comment only: each reply keeps its own range on the old text (a reply has its parent's range since rdocx 0.17.0), a reply id is refused ("not a movable root"), and removing the old text then removes the replies with it | `docx_ops.move_thread_to_text(doc, id, anchor, occurrence=0)` moves the root, then adds each reply again under it with its author, initials, date, text and resolved flag (a reply's id can change: it returns {old id: new id}). From the command line, run it through `$R/python` |
| Spanning over a formatted empty cell [docx-grid-span-refuses-formatted-empty-cell] | `table.set_cell_grid_span(r, c, n)` raises `horizontal merge cannot consume a nonempty cell` when a cell it would consume is empty but its paragraph carries properties (a right-aligned totals column, a cell made by `insert_column` next to one) | `cell.merge(other)` consumes such a cell and keeps no blank paragraph: merge the cells left to right instead |
| Replacing through Python [docx-python-replace-slower-than-cli] | `Document.open` + `try_replace_text` + `save` (and `docx_ops.replace_batch`, which uses them) take about twice `rdocx replace` on the same file, 3 to 4 times on reports with tables, notes, comments and tracked changes: 0.8 + 1.6 s for 111 replacements in a 52-page report, against 0.67 s for the CLI. The binding builds its index of the document's elements when it opens it and again after the replacement | For replacements alone in a large file, run `rdocx replace F --map pairs.json -o OUT` (each pair with its expected count, all or nothing), then check the output as usual. Keep the binding when other edits follow in the same document |

Edit costs of a large document, accepted: an edit still costs a little more in a larger body. At 1000 paragraphs as
Google Docs writes them (1.7 MB of XML, measured on Linux), `insert_paragraph`, `add_paragraph` and `remove_content`
take 8 to 11 ms a call, `p.text` 27 ms, `run.text` 0.14 s, the replacements (`p.replace_text`, `cell.replace_text`,
`replace_text_at`) 0.17 to 0.18 s and `add_picture` 0.24 s (in a cell too): one replacement in each of those
paragraphs takes about 3 minutes, a `run.text` on each of their 6000 runs about 14. Calls made once or a few times
per document cost more: `move_comment` 2 s, `add_heading_hyperlink` 1.6 s, `move_content` and internal hyperlinks
about 1 s, the section and page calls (`remove_section`, `create_section_story`, `link_section_story`,
`unlink_section_story`, `is_linked_to_previous`, `set_page_color`) 0.5 to 0.7 s, accepting or rejecting a revision
by id 0.6 s, `insert_document` 0.23 s, `set_header` and `set_footer` 0.16 s. A large `styles.xml` costs nothing per
edit. A cell property setter (`width`, `vertical_alignment`, `shading`, `set_border`, `set_margins`) costs the same
whatever the body, `cell.text` what `p.text` costs, and a setter through a cell's paragraph or run the same in any
row, but in a cell inside a content control (`references/python-api.md`, Handles). Make many replacements in one batch
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
