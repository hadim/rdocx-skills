# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Edits of a large document [docx-edit-cost-grows-with-body] | 75 of the 252 calls `tests/test_docx_scaling.py` times still cost a time that grows with the body's XML: content controls (insert, set, move, clone, remove), footnotes and endnotes, sections and their stories, `set_header` and `set_footer`, pictures and their sizes, moving, copying or importing content (`move_content`, `clone_content`, `pop_content`, `insert_content`, `insert_document`, `insert_html`, `import_fragment`), story text and hyperlinks, moving or removing a comment, accepting or rejecting a revision, list items, `run.text`, `cell.text`, replacements in a paragraph or a cell, `replace_xml`, table rows and columns, merging and splitting cells. In that test's body at 1000 paragraphs, as Google Docs writes them (1.7 MB of XML), the ones measured take 0.4 to 1.1 s a call (`run.text` 0.5 s, `insert_content_control` 0.8 s, `add_picture` 1.1 s). Flat in the test and cheap there: adding, inserting and removing paragraphs, headings, page breaks and tables (`insert_paragraph` and `remove_content` about 2 ms a call, `p.text` 16 ms), `p.add_run`, `p.add_hyperlink`, adding, answering and resolving comments, table and cell formatting, run, font and paragraph formatting (`run.bold`, `font.*`, `paragraph_format.*`, `p.style`, `p.alignment`), page setup (`doc.section(i)` sizes and margins, `update_section`), watermarks, fonts, numbering definitions, styles, document defaults, bookmarks, document properties, reading `p.xml` | A few such edits cost a second or two. For many, make them in a document built aside, which is small, then one `insert_document` (0.6 s at 1000 paragraphs): a new `rdocx.Document()`, or a copy of the target emptied by one `remove_content_range(start, end)` (`rdocx.Document.from_bytes(doc.to_bytes())`) to start from its styles. |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
