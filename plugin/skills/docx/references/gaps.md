# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Edits of a large document [docx-edit-cost-grows-with-body] | Most edits still cost a time that grows with the body's XML, if much less than on rdocx 0.16.3. At 1000 paragraphs as Google Docs writes them (1.7 MB of XML, measured on Linux), the 14 calls `tests/test_docx_scaling.py` still counts as growing: `move_comment` 1.4 s, replacements (`p.replace_text`, `cell.replace_text`, `replace_text_at`) about 1 s, `move_content`, `add_picture` (also in a cell) and internal hyperlinks 0.65 to 0.75 s, accepting or rejecting a revision by id 0.55 s, `set_header` and `set_footer` 0.37 s, story text and hyperlinks 0.15 to 0.22 s. The others grow less, some costing as much at that size: `add_heading_hyperlink` 0.46 s, `insert_document`, `insert_html` and cloning a content control 0.23 to 0.25 s, `clone_content`, removing a comment, inserting or moving a content control, `replace_xml` or a picture size 0.13 to 0.18 s, `cell.text`, a footnote or a table row about 80 ms, `run.text` 75 ms, `p.text` 20 ms. Only `insert_paragraph`, `add_paragraph` and `remove_content` stay at about 2 ms, so a loop over every run of such a document still takes minutes | For many edits of a large document, make them in a document built aside, which is small, then one `insert_document`: a new `rdocx.Document()`, or a copy of the target emptied by one `remove_content_range(start, end)` (`rdocx.Document.from_bytes(doc.to_bytes())`) to start from its styles. Where the runs' formatting does not matter, one `p.text` per paragraph costs less than a `run.text` on each run. |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
