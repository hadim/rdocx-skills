# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`; `doc.part_xml(name)` reads any XML part, the relationships and
`[Content_Types].xml` included: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Edits of a large document [docx-edit-cost-grows-with-body] | 145 of the 252 calls `tests/test_docx_scaling.py` times cost a time that grows with the body's XML, whatever the element edited: adding, inserting, moving or removing content (paragraphs, headings, tables, pictures, charts, list items, notes, comments, sections), text (`p.text`, `run.text`, `cell.text`, replacements), `replace_xml`, table structure and cell edits, content controls, hyperlinks, revisions, headers, footers, watermarks, fonts, numbering, and reading `p.xml`. In a body of 1000 paragraphs as Google Docs writes them (2.4 MB of XML), `insert_paragraph` and `remove_content` take about 0.5 s a call, `p.xml` 0.15 s. Constant: run, font and paragraph formatting (`run.bold`, `font.*`, `paragraph_format.*`, `p.style`, `p.alignment`), page setup (`doc.section(i)` sizes and margins, `update_section`), styles, document defaults, bookmarks, document properties | Edit the structure in few calls: one `remove_content_range(start, end)` for a run of blocks (70 blocks of that body: 1.8 s, against 35 s one by one), and the new content built in a document aside, then one `insert_document` (70 paragraphs: 3 s against 40 s). Build that document as a new `rdocx.Document()`, or in a copy of the target emptied by one `remove_content_range` (`rdocx.Document.from_bytes(doc.to_bytes())`) to start from its styles. To read the XML of many paragraphs, parse `doc.part_xml("word/document.xml")` once with `xml.etree.ElementTree`: 0.5 s for that body, against 0.15 s a paragraph |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
