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
| body index | `find_content_index(handle or text)`, `find_content_indices(text)`, `insert_paragraph`, `remove_content`, `pop_content`, `insert_content`, `clone_content` and `move_content` destinations, `split_run(i, ...)`, `RunPosition.body_index`, `LayoutFragment.body_index`, `StoryItem.direct_body_index`, `rdocx text --json` `body_index`, `rdocx comment add --start-paragraph` | top-level children of the body: paragraphs, tables, content-control blocks (the TOC is one block) |
| flow index | `doc.paragraphs[i]` | every paragraph outside tables, those inside content-control blocks (TOC entries) included, and the cell paragraphs of a table inside such a block |
| story path | `doc.story_items`, `StoryItem.index_path`, `Hyperlink.index_path`, `set_story_text(item, ...)` | per story (body, table cell, text box, header, footer, footnote, ...): like the body index plus items for pictures, content controls and fields |

- `find_content_index(handle)` raises `ValueError` for a paragraph that is not a direct child of the body
  (a TOC entry, a paragraph in a content-control block or a table cell).
- `find_content_indices(text)` returns every top-level block whose text contains `text`, the TOC block
  included for any heading (never delete by it); it does not search tables. `find_content_index(text)`
  returns the first, or raises `ValueError`.
- Run indices of `RunPosition` and `split_run` count the runs `Paragraph.runs` lists, those inside an inline
  content control or a tracked insertion included. `Paragraph.runs` does not hold the text of a simple
  field, a smart tag or a custom XML element, which `Paragraph.text` holds: a character offset in `p.text`
  is not a run offset. `docx_ops.locate` returns run
  offsets, and `docx_ops.comment_on_text` finds a text and anchors on it.

## Handles

Every handle is checked: `Paragraph`, `Run`, `Table`, `Row`, `Cell` and `StoryItem` (and the
`StoryRunRange` built on one) raise `StaleElementError` once the document has changed structure. Only
formatting setters (font, paragraph format, style, alignment, numbering, widths) and `Run.text` keep handles
valid. Everything else invalidates every handle: insert, remove, clone, move, pop, split, `add_*`,
`set_story_text`, `paragraph.text = ...`, `cell.text = ...`, `try_replace_text`, `replace_all_regex`,
`clone_row`, `remove_row`, `add_comment`. Re-navigate from `doc` after each such call (`doc.paragraphs[i].runs[j]`,
`doc.story_items`); never keep `p = doc.paragraphs[i]` across edits.

## Document

| Member | Notes |
|---|---|
| `Document()`, `Document.open(path)`, `Document.from_bytes(b)`, `doc.to_bytes()`, `doc.save(path)` | a new document has Word's usual styles (Normal, Title, Subtitle, Heading 1 to 9, No Spacing, Quote, List Paragraph, Caption, Table Grid). `save` writes a temporary file and renames it (`docx_ops.save_atomic` also refuses the input), and writes the package class its extension names: a .dotx saved as .docx becomes a document |
| `paragraphs`, `tables`, `sections`, `styles`, `stories`, `story_items`, `header_footer_variants`, `hyperlinks`, `comments`, `revisions` | read-only collections (sections and styles are snapshots; `revisions` covers every story, `r.story.kind`) |
| `core_properties` → `CoreProperties` | `title`, `author`, `subject`, `keywords`, `category`, `comments`, `last_modified_by`, `content_status`, `language`, `identifier`, `version` (strings), `created`, `modified`, `last_printed` (`datetime` or None), `revision` (int): read and set |
| `add_style(name, style_type="paragraph", *, style_id=None, based_on=None, next_style=None, font_name=, font_size=, bold=, italic=, color=, space_before=, space_after=, left_indent=, right_indent=, first_line_indent=)` → `Style` | as python-docx's `styles.add_style`: unless given, the id keeps the letters, digits and hyphens of the name (`"Note box"` → `Notebox`). `based_on` and `next_style` take an id or a name. Lengths and `font_size` in EMU (`rdocx.Pt(11)`). `KeyError` for an unknown base, `ValueError` for a duplicate |
| `remove_style(style)` → bool, `set_default_style(style)` | |
| `add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.", start=1, left_indent=, hanging_indent=), ...])` → definition id, `add_numbering_instance(definition_id)` → `num_id`, `link_style_to_numbering(style, num_id, level)` | a list: define its levels once, make an instance, then `paragraph.numbering = (num_id, level)` or link a style to it |
| `add_paragraph(text)`, `insert_paragraph(bi, text)`, `add_table(rows, cols)` | append / insert with Normal formatting |
| `add_picture(data, filename, width=None, height=None, *, after=None)` | inline picture in a new paragraph; width and height together or neither; returns the `StoryItem` |
| `clone_content(handle, bi)` | copies a block to body index `bi`, with its formatting, fields and bookmarks (a copied bookmark is renamed, `MailMerge1`...); comment anchors are not copied |
| `move_content(handle, bi)`, `pop_content(bi)` → `ContentFragment`, `insert_content(bi, fragment)`, `remove_content(bi)` | `bi` of `move_content` is counted before the move (moving the 2nd of ABCDE to 3 gives ACBDE) |
| `set_story_text(item, text)` | replaces a paragraph's text, keeping the first run's formatting (any story) |
| `split_run(body_index, run_index, offset)` | splits one run in two at a character offset; body paragraphs only |
| `update_section(i, *, margin_top=, orientation=, page_width=, ...)` → `Section` | writes a section's page setup; `doc.sections` stays a snapshot |
| `try_replace_text(old, new, *, expect=None)` → int, `replace_all_regex([(pattern, repl), ...])` → int | literal and regex replacement across runs, in the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes (a Word text box once), headers and footers with their tables (once per variant part), footnotes and endnotes. A match across the edge of a content control, an insertion or a simple field is not replaced. With `expect=N`, a different count raises and changes nothing; `docx_ops.replace_batch` also refuses text left out of reach |
| `set_header(text)`, `set_footer(text)`, `add_hyperlink_to_story(story, text, url)` | `set_header` / `set_footer` replace the default story's content (fields included) |
| `set_hyperlink_url(link, url)`, `remove_hyperlink(link)` | `link` from `doc.hyperlinks`. Removal keeps the text. Retargeting keeps the other entries of `doc.hyperlinks` valid. After a removal or an added link, the older entries of that story raise `RdocxError`: re-fetch `doc.hyperlinks` |
| `image_data(rid)`, `replace_image(rid, bytes)`, `replace_image_for_story(story, rid, bytes)`, `set_picture_size(rid, width, height)` → count | the relationship id is the `r:embed` of the picture (`StoryItem.xml` of a `drawing` item). `replace_image` keeps the old extent, `set_picture_size` resizes every body picture of that relationship (EMU) |
| `add_comment(RunRange or StoryRunRange, *, author, text, initials=None, date=None)` → id | `RunRange(start=RunPosition(body_index=, run_index=), end=...)`, end exclusive, body paragraphs; `StoryRunRange(start=StoryRunPosition(item=, run_index=), end=...)` for a table cell or another story's paragraph item. No `date`, no date in the file: pass an RFC 3339 string. Ids stay stable across saves |
| `reply_to(parent_id, *, author, text, date=None)` → id, `resolve_comment(id, *, resolved=True)`, `remove_comment(id)` | |
| `accept_all()`, `reject_all()`, `accept_revision_id(id)`, `reject_revision_id(id)`, `accept_revisions_by_author(a)`, `reject_revisions_by_author(a)`, `accept_revisions_in_date_range(*, start, end)`, `reject_revisions_in_date_range(*, start, end)` | every supported story; dates as RFC 3339 strings |
| `compare(edited, author, timestamp, *, granularity="run", ignore_comments=False, ...)` → diagnostics | turns `doc` into the redline of `doc` → `edited`; `granularity="word"` marks only the changed words, `ignore_comments=True` keeps `doc`'s comments and compares the rest |
| `rebuild_toc()` → `TocRebuildReport` (`entry_count`, `bookmark_count`, `diagnostics`) | entries of numbered headings need a left tab stop in the TOC styles (gap toc-numbered-entries) |
| `update_page_fields()` → int, `update_layout_backed_fields()` → report (`page_fields`, `num_pages_fields`, `page_reference_fields`, `updated_count`, `diagnostics`), `update_fields(*, now=datetime, file_name=, file_path=, merge_fields=, ...)` → count, `update_fields_on_open` (get/set) | caches field results from rdocx's pagination |
| `layout()` → tuple of `LayoutFragment` (`body_index`, `physical_page`, `displayed_page`, `bounds.x/.y/.width/.height` in points), `layout_page(i)` → `LayoutPage` (`page_number`, `displayed_page_number`, `width`, `height`) | rdocx's pagination (gaps line-gap, picture-line-spacing) |
| `to_pdf()` → bytes, `render_pages(*, dpi=150, format="png", quality=90, transparent=False, pages=None)` → list of bytes, `render_page_to_png(i, dpi)`, `render_all_pages(dpi)` | `pages` zero-based |

## Paragraph, Run, Font, ParagraphFormat

- `Paragraph`: `text` (the accepted view: tracked insertions in, deletions out. Setting it leaves one run
  without direct formatting, keeps the paragraph style, format and comment ranges, and refuses a
  paragraph holding one end of a field that spans paragraphs, such as a TOC), `runs`, `style` (reads the
  style id. Set an id or a name the document defines, as python-docx does: `KeyError` when no paragraph
  style has it, `ValueError` for a character or table style), `alignment` (`WD_ALIGN_PARAGRAPH`),
  `numbering` (`(num_id, level)` or None, settable, unchecked), `paragraph_format`, `add_run(text)`,
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
`vertical_alignment` (`WD_CELL_VERTICAL_ALIGNMENT`), `shading` (hex fill, settable), `grid_span`,
`vertical_merge`. On the table: `set_cell_grid_span(r, c, n)` (consumes empty cells only, and invalidates
the table handle), `set_cell_vertical_merge(r, c, "restart" | "continue" | None)`, `set_borders(style, *,
size, color)`, `set_border(edge, ...)`, `set_cell_margins(...)`, `set_column_width(c, w)` (the grid);
`Row.height` and `height_rule` (settable).

## Stories other than the body

`doc.story_items` lists the paragraphs, tables, pictures, content controls and fields of every story, with
`item.story.kind` in `body`, `table_cell`, `text_box`, `header`, `footer`, `footnote`, `endnote`, ... and
`item.story.part_name`. A header or footer has one story per variant part, and a text box as Word writes
it (`mc:AlternateContent`, DrawingML with a VML copy) is listed once. `doc.paragraphs` holds the body only,
and `rdocx text` and `text --json` (`stories`) print the others after it. `docx_ops.story_paragraphs(doc)`,
`all_text(doc)` and `count(doc, text)` read the package XML: every paragraph of every story, in the accepted
view, with its part and style id.

## Read-only records

`Section` (orientation, page size, margins, gutter, columns, header and footer distances, `different_first_page`,
`break_type`, `page_number_start`); `Style` (`style_id`, `name`, `style_type`, `based_on`, `linked_style`,
`next_style`, flags); `HeaderFooterVariant` (`section_index`, `kind` header/footer, `variant`
default/first/even, `story`, `inherited`, `source_section`); `Hyperlink` (`url`, `anchor`, `text`,
`index_path`, `relationship_id`, `story`); `Comment` (`id`, `author`, `initials`, `date`, `text`,
`parent_id`, `resolved`); `Revision` (`id`, `kind`, `author`, `timestamp`); `StoryItem` (`kind` paragraph,
table, drawing, content_control, field, preserved_node; `index_path`, `direct_body_index`, `text`, `xml`,
`story`); `Story` (`kind`, `part_name`, `owner_index`).
