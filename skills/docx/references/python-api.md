# rdocx Python API (pinned build)

`import rdocx` with the pinned Python (`$R/python`). The surface looks like python-docx where it overlaps,
with document-level methods for everything python-docx lacks (comments, revisions, comparison, layout,
rendering). The modules have no docstrings: exact signatures are in the type stubs, `rdocx/_rdocx.pyi` next
to the installed module (`$R/python -c "import rdocx, os; print(os.path.dirname(rdocx.__file__))"`).

Lengths are EMU integers; build them with `rdocx.Pt(12)`, `rdocx.Inches(1)`, `rdocx.Cm(2)`, `rdocx.Mm(5)`,
`rdocx.Emu(n)` (each has `.pt`, `.inches`, `.cm`, `.mm`, `.emu`, `.twips`). Errors: `rdocx.RdocxError`
(base), `XmlError`, `PackageError`, `LayoutError`, `StaleElementError`; lookups by text raise `ValueError`.

## Three index spaces

| Space | Used by | Counts |
|---|---|---|
| body index | `find_content_index(handle or text)`, `find_content_indices(text)`, `insert_paragraph`, `remove_content`, `pop_content`, `insert_content`, `clone_content` and `move_content` destinations, `RunPosition.body_index`, `LayoutFragment.body_index`, `StoryItem.direct_body_index`, `rdocx text --json` `body_index`, `rdocx comment add --start-paragraph` | top-level children of the body: paragraphs, tables, content-control blocks (the TOC is one block) |
| flow index | `doc.paragraphs[i]`, `split_run(i, ...)` | every paragraph outside tables, those inside content-control blocks (TOC entries) included, and the cell paragraphs of a table inside such a block |
| story path | `doc.story_items`, `StoryItem.index_path`, `Hyperlink.index_path`, `set_story_text(item, ...)` | per story (body, table cell, text box, header, footer, footnote, ...): like the body index plus items for pictures, content controls and fields |

- `find_content_index(handle)` raises `ValueError` for a paragraph that is not a direct child of the body
  (a TOC entry, a paragraph in a content-control block or a table cell).
- `find_content_indices(text)` returns every top-level block whose text contains `text`, the TOC block
  included for any heading (never delete by it); it does not search tables. `find_content_index(text)`
  returns the first, or raises `ValueError`.
- **Run indices of `RunPosition` skip runs inside an inline content control or a tracked insertion**, while
  `Paragraph.runs` counts them (gap comment-runposition-sdt). Use `docx_ops.comment_on_text`, which checks
  the anchored text.

## Handles

Every handle is checked: `Paragraph`, `Run`, `Table`, `Row`, `Cell` and `StoryItem` (and the
`StoryRunRange` built on one) raise `StaleElementError` once the document has changed structure. Only
formatting setters (font, paragraph format, style, alignment, numbering, widths) and `Run.text` keep handles
valid. Everything else invalidates every handle: insert, remove, clone, move, pop, split, `add_*`,
`set_story_text`, `cell.text = ...`, `try_replace_text`, `replace_all_regex`, `clone_row`, `remove_row`,
`add_comment`. Re-navigate from `doc` after each such call (`doc.paragraphs[i].runs[j]`,
`doc.story_items`); never keep `p = doc.paragraphs[i]` across edits.

## Document

| Member | Notes |
|---|---|
| `Document()`, `Document.open(path)`, `Document.from_bytes(b)`, `doc.to_bytes()`, `doc.save(path)` | new documents have the Normal and Heading1 styles only (gap new-document-styles); `save` writes in place (use `docx_ops.save_atomic`); a .dotx saved as .docx keeps the template content type (gap template-save-as-document): start from a copy of a .docx |
| `paragraphs`, `tables`, `sections`, `styles`, `stories`, `story_items`, `header_footer_variants`, `hyperlinks`, `comments`, `revisions` | read-only collections (sections and styles are snapshots) |
| `add_paragraph(text)`, `insert_paragraph(bi, text)`, `add_table(rows, cols)` | append / insert with Normal formatting |
| `add_picture(data, filename, width=None, height=None, *, after=None)` | inline picture in a new paragraph; width and height together or neither; returns the `StoryItem` |
| `clone_content(handle, bi)` | copies a block to body index `bi`, with its formatting, fields and bookmarks (a copied bookmark is renamed, `MailMerge1`...); comment anchors are not copied |
| `move_content(handle, bi)`, `pop_content(bi)` → `ContentFragment`, `insert_content(bi, fragment)`, `remove_content(bi)` | `bi` of `move_content` is counted before the move (moving the 2nd of ABCDE to 3 gives ACBDE) |
| `set_story_text(item, text)` | replaces a paragraph's text, keeping the first run's formatting |
| `split_run(flow_index, run_index, offset)` | splits one run in two at a character offset; flow index (gap split-run-index); body paragraphs only |
| `try_replace_text(old, new)` → int, `replace_all_regex([(pattern, repl), ...])` → int | literal and regex replacement across runs, in the body, its tables and text boxes (a Word text box counts twice), headers and footers (once per variant part); not inside content controls, tracked insertions, footnotes, endnotes, header and footer tables, simple fields or smart tags (gaps); no expected count (use `docx_ops.replace_batch`) |
| `set_header(text)`, `set_footer(text)`, `add_hyperlink_to_story(story, text, url)` | `set_header` / `set_footer` replace the default story's content (fields included) |
| `image_data(rid)`, `replace_image(rid, bytes)`, `replace_image_for_story(story, rid, bytes)` | the relationship id is the `r:embed` of the picture (`StoryItem.xml` of a `drawing` item) |
| `add_comment(RunRange or StoryRunRange, *, author, text, initials=None, date=None)` → id | `RunRange(start=RunPosition(body_index=, run_index=), end=...)`, end exclusive, body paragraphs; `StoryRunRange(start=StoryRunPosition(item=, run_index=), end=...)` for a table cell or another story's paragraph item. No `date`, no date in the file: pass an RFC 3339 string. Ids stay stable across saves |
| `reply_to(parent_id, *, author, text, date=None)` → id, `resolve_comment(id, *, resolved=True)`, `remove_comment(id)` | |
| `accept_all()`, `reject_all()`, `accept_revision_id(id)`, `reject_revision_id(id)`, `accept_revisions_by_author(a)`, `reject_revisions_by_author(a)`, `accept_revisions_in_date_range(*, start, end)`, `reject_revisions_in_date_range(*, start, end)` | every supported story; dates as RFC 3339 strings |
| `compare(edited, author, timestamp)` | turns `doc` into the redline of `doc` → `edited` |
| `rebuild_toc()` → `TocRebuildReport` (`entry_count`, `bookmark_count`, `diagnostics`) | gap toc-rsid-field-runs: use `docx_ops.rebuild_toc(doc)`; entries of numbered headings need a left tab stop in the TOC styles (gap toc-numbered-entries) |
| `update_page_fields()` → int, `update_layout_backed_fields()` → report (`page_fields`, `num_pages_fields`, `page_reference_fields`, `updated_count`, `diagnostics`), `update_fields(*, now=datetime, file_name=, file_path=, merge_fields=, ...)` → count, `update_fields_on_open` (get/set) | caches field results from rdocx's pagination |
| `layout()` → tuple of `LayoutFragment` (`body_index`, `physical_page`, `displayed_page`, `bounds.x/.y/.width/.height` in points), `layout_page(i)` → `LayoutPage` (`page_number`, `displayed_page_number`, `width`, `height`) | rdocx's pagination (gaps line-gap, picture-line-spacing) |
| `to_pdf()` → bytes, `render_pages(*, dpi=150, format="png", quality=90, transparent=False, pages=None)` → list of bytes, `render_page_to_png(i, dpi)`, `render_all_pages(dpi)` | `pages` zero-based |

No core properties (title, author, dates) in either direction (gap docx-core-properties).

## Paragraph, Run, Font, ParagraphFormat

- `Paragraph`: `text` (read-only; the accepted view: tracked insertions in, deletions out), `runs`, `style`
  (style **id**, settable; an unknown id or a style name is accepted silently, gap style-id-unchecked:
  check it against `{s.style_id for s in doc.styles}`), `alignment` (`WD_ALIGN_PARAGRAPH`), `numbering`
  (`(num_id, level)` or None, settable, unchecked too), `paragraph_format`, `add_run(text)`,
  `add_hyperlink(text, url)` (appends at the end).
- `ParagraphFormat`: `space_before`, `space_after`, `line_spacing` (float for multiples, length for exact),
  `left_indent`, `right_indent`, `first_line_indent`, `keep_with_next`, `keep_together`,
  `page_break_before`, `widow_control`, `alignment`. None means inherited.
- `Run`: `text` (settable), `font`, `style_id`.
- `Font`: `name`, `size` (length), `bold`, `italic`, `underline` (`WD_UNDERLINE` or bool), `strike`,
  `color` (set an `rdocx.RGBColor(r, g, b)`; reads back an `RGBColor`, a tuple: compare `str(font.color) ==
  "7B1E3A"`; a hex string is refused), `highlight` (Word colour name: "yellow", "green", ...), `shading`
  (hex fill). None means inherited.

## Tables

`Table`: `rows`, `cell(r, c)`, `clone_row(index, at=None)` (keeps cell and run formatting),
`remove_row(index)`, `style` (style id, unchecked), `alignment` (`WD_TABLE_ALIGNMENT`), `width` (settable).
`Row.cells` (a merged cell appears once: use `cells[-1]` for the last column); `Cell`: `text` (settable),
`paragraphs`, `add_paragraph(text)`, `width` (settable: writes the cell width, not the table grid),
`vertical_alignment` (`WD_CELL_VERTICAL_ALIGNMENT`). No merge, borders, shading, grid or row height from
Python (gap docx-python-tables).

## Stories other than the body

`doc.story_items` lists the paragraphs, tables, pictures, content controls and fields of every story, with
`item.story.kind` in `body`, `table_cell`, `text_box`, `header`, `footer`, `footnote`, `endnote`, ... and
`item.story.part_name`; a header or footer has one story per variant part. It misses text boxes as Word
writes them (`mc:AlternateContent`, gap textbox-alternate-content), and the CLI views and `doc.paragraphs`
leave out every story but the body. `docx_ops.story_paragraphs(doc)`, `all_text(doc)` and `count(doc, text)`
read the package XML instead: every paragraph of every story, in the accepted view, a Word text box once.

## Read-only records

`Section` (orientation, page size, margins, gutter, columns, header and footer distances, `different_first_page`,
`break_type`, `page_number_start`); `Style` (`style_id`, `name`, `style_type`, `based_on`, `linked_style`,
`next_style`, flags); `HeaderFooterVariant` (`section_index`, `kind` header/footer, `variant`
default/first/even, `story`, `inherited`, `source_section`); `Hyperlink` (`url`, `anchor`, `text`,
`index_path`, `relationship_id`, `story`); `Comment` (`id`, `author`, `initials`, `date`, `text`,
`parent_id`, `resolved`); `Revision` (`id`, `kind`, `author`, `timestamp`); `StoryItem` (`kind` paragraph,
table, drawing, content_control, field, preserved_node; `index_path`, `direct_body_index`, `text`, `xml`,
`story`); `Story` (`kind`, `part_name`, `owner_index`).
