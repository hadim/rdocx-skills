# rdocx Python API (pinned build)

`import rdocx` with the pinned Python (`$R/python`). The surface looks like python-docx where it overlaps,
with document-level methods for everything python-docx lacks (comments, revisions, comparison, layout,
rendering). The modules have no docstrings: exact signatures are in the type stubs, `rdocx/_rdocx.pyi` next
to the installed module (`$R/python -c "import rdocx, os; print(os.path.dirname(rdocx.__file__))"`).

Lengths are EMU integers; build them with `rdocx.Pt(12)`, `rdocx.Inches(1)`, `rdocx.Cm(2)`, `rdocx.Mm(5)`,
`rdocx.Emu(n)` (each has `.pt`, `.inches`, `.cm`, `.mm`, `.emu`, `.twips`); widths, margins and font sizes
read back as `rdocx.Length`, the int subclass they share. Colours: `rdocx.RGBColor(r, g, b)` or
`rdocx.RGBColor.from_string("7B1E3A")`. Errors: `rdocx.RdocxError`
(base), `XmlError`, `PackageError`, `LayoutError`, `StaleElementError`, `ReplacementCountError` (an
`expect` count not met); lookups by text raise `ValueError`.

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
`paragraph.replace_text`, `cell.replace_text`, `replace_text_at`, `clone_row`, `remove_row`, `add_comment`. Re-navigate from `doc` after each such call (`doc.paragraphs[i].runs[j]`,
`doc.story_items`); never keep `p = doc.paragraphs[i]` across edits.

## Document

| Member | Notes |
|---|---|
| `Document()`, `Document.open(path)`, `Document.from_bytes(b)`, `doc.to_bytes()`, `doc.save(path)` | a new document has Word's usual styles (Normal, Title, Subtitle, Heading 1 to 9, No Spacing, Quote, List Paragraph, Caption, Table Grid). `save` writes a temporary file and renames it (`docx_ops.save_atomic` also refuses the input), and writes the package class its extension names: a .dotx saved as .docx becomes a document |
| `paragraphs`, `tables`, `sections`, `styles`, `stories`, `story_items`, `header_footer_variants`, `hyperlinks`, `comments`, `revisions` | read-only collections (sections and styles are snapshots; `revisions` covers every story, `r.story.kind`) |
| `core_properties` → `CoreProperties` | `title`, `author`, `subject`, `keywords`, `category`, `comments`, `last_modified_by`, `content_status`, `language`, `identifier`, `version` (strings), `created`, `modified`, `last_printed` (`datetime` or None), `revision` (int): read and set |
| `add_style(name, style_type="paragraph", *, style_id=None, based_on=None, next_style=None, font_name=, font_size=, bold=, italic=, color=, space_before=, space_after=, left_indent=, right_indent=, first_line_indent=)` → `Style` | as python-docx's `styles.add_style`: unless given, the id keeps the letters, digits and hyphens of the name (`"Note box"` → `Notebox`). `based_on` and `next_style` take an id or a name. Lengths and `font_size` in EMU (`rdocx.Pt(11)`). `KeyError` for an unknown base, `ValueError` for a duplicate |
| `set_style(style, *, based_on=, next_style=, font_name=, font_size=, bold=, italic=, color=, space_before=, space_after=, left_indent=, right_indent=, first_line_indent=)` → `Style` | changes an existing style, chosen by id or name, as `add_style` takes them; what it is not given keeps its value. It cannot remove a theme font or a theme colour the style already has, which Word keeps using over `font_name` / `color`. `KeyError` for an unknown style, and on any error the document is left unchanged |
| `remove_style(style)` → bool, `set_default_style(style)` | |
| `add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.", start=1, left_indent=, hanging_indent=), ...])` → definition id, `add_numbering_instance(definition_id)` → `num_id`, `link_style_to_numbering(style, num_id, level)` | a list: define its levels once, make an instance, then `paragraph.numbering = (num_id, level)` or link a style to it |
| `add_paragraph(text)`, `insert_paragraph(bi, text)`, `add_table(rows, cols)`, `insert_table(bi, rows, cols)` → `Table` | append / insert with Normal formatting |
| `add_picture(data, filename, width=None, height=None, *, after=None)` | inline picture in a new paragraph; width and height together or neither; returns the `StoryItem` |
| `clone_content(handle, bi)` | copies a block to body index `bi`, with its formatting, fields and bookmarks (a copied bookmark is renamed, `MailMerge1`...); comment anchors are not copied |
| `move_content(source, destination)`, `pop_content(bi)` → `ContentFragment`, `insert_content(bi, fragment)`, `remove_content(bi)` | `destination` of `move_content` is counted before the move (moving the 2nd of ABCDE to 3 gives ACBDE). Removing content never leaves a comment without an anchor: a comment it covers whole goes with its replies; a removal that would cut a comment in part, and `pop_content` of content that carries a comment (fragments do not own threads), raise `RdocxError` and change nothing. Move the thread first, or use `move_content`, which keeps it on the moved block |
| `set_story_text(item, text)` | replaces a paragraph's text, keeping the first run's formatting (any story) |
| `split_run(body_index, run_index, character_offset)` | splits one run in two at a character offset; body paragraphs only (a table cell paragraph raises `ValueError`) |
| `update_section(i, *, margin_top=, orientation=, page_width=, ...)` → `Section` | writes a section's page setup; `doc.sections` stays a snapshot |
| `insert_section(i)`, `remove_section(i)` | `insert_section` adds an empty section at section index `i` (an empty paragraph carries its break: 0 before the first section, `len(doc.sections)` after the last one's content); `remove_section` removes that break, the paragraphs stay |
| `create_section_story(i, kind, variant)` → `Story`, `unlink_section_story(i, kind, variant)` → `Story`, `link_section_story(i, kind, variant, story)` → `Story` | `kind` `"header"` / `"footer"`, `variant` `"default"` / `"first"` / `"even"`. `create_section_story` gives section `i` a new, empty header or footer (`"first"` turns on `different_first_page`), which the next sections inherit (`doc.header_footer_variants`); fill it with `insert_content(story, fragment)` (a fragment from `pop_content`). `unlink_section_story` gives a section that inherits one its own copy, to edit apart; `link_section_story` makes it use `story` again |
| `try_replace_text(old, new, *, expect=None)` → int, `replace_all_regex([(pattern, repl), ...])` → int | literal and regex replacement across runs, in the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes (a Word text box once), headers and footers with their tables (once per variant part), footnotes and endnotes. A match across the edge of a content control, an insertion or a simple field is not replaced. With `expect=N`, a different count raises `ReplacementCountError` and changes nothing; `docx_ops.replace_batch` also refuses text left out of reach |
| `replace_all([(old, new), (old, new, expect), ...])` → tuple of counts | several literal replacements in one call, all or nothing: a pair whose `expect` is not met raises `ReplacementCountError` (naming the pair) and changes nothing |
| `replace_text_at(item, old, new, *, expect=None)` → int | the same replacement in one `StoryItem` only (`doc.story_items`): a body paragraph or table, a table cell's paragraph, a header, footer or footnote paragraph. A comment paragraph is its own text (not the anchor). A text box: gap replace-text-at-text-box-copy, use the document-wide call, which edits both copies of a Word text box. For one body paragraph or cell, `Paragraph.replace_text` and `Cell.replace_text` |
| `bookmarks` → tuple of `Bookmark` (`id`, `name`, `text`, `range`, `direct_range`, `issue`), `add_bookmark(name, range)` → id | `add_bookmark` takes a body `RunRange`; `direct_range` gives it back that way (None for a bookmark in a table cell or a content-control block), `range` counts paragraphs through tables; `issue` describes a broken marker, None otherwise |
| `set_header(text)`, `set_footer(text)`, `add_hyperlink_to_story(story, text, url)` | `set_header` / `set_footer` replace the default story's content (fields included); a comment anchored only in the replaced header or footer goes with it, replies included |
| `set_hyperlink_url(hyperlink, url)`, `remove_hyperlink(hyperlink)` | `hyperlink` from `doc.hyperlinks`. Removal keeps the text. Retargeting keeps the other entries of `doc.hyperlinks` valid. After a removal or an added link, the older entries of that story raise `RdocxError`: re-fetch `doc.hyperlinks` |
| `image_data(rid)`, `replace_image(rid, bytes)`, `replace_image_for_story(story, rid, bytes)`, `set_picture_size(rid, width, height)` → count | the relationship id is the `r:embed` of the picture (`StoryItem.xml` of a `drawing` item). `replace_image` keeps the old extent, `set_picture_size` resizes every body picture of that relationship (EMU) |
| `add_comment(RunRange or StoryRunRange, *, author, text, initials=None, date=None)` → id | `RunRange(start=RunPosition(body_index=, run_index=), end=...)`, end exclusive, body paragraphs; `StoryRunRange(start=StoryRunPosition(item=, run_index=), end=...)` for a table cell or another story's paragraph item, or `StoryRunPosition(paragraph=p, run_index=)` from a `Paragraph` handle, a paragraph of a content-control block included (one in a nested control or in a table inside the block raises `IndexError`). No `date`, no date in the file: pass an RFC 3339 string. Ids stay stable across saves |
| `add_comment_on_text(anchor, *, author, text, occurrence=0, initials=None, date=None)` → id | anchors on exactly `anchor`, splitting runs with their format kept, table cells included: `occurrence` counts from 0 over the main story in document order, body and table cells together; an out-of-range occurrence raises `RdocxError`. `docx_ops.comment_on_text(..., in_tables=True)` counts cells only and checks the result |
| `reply_to(parent_id, *, author, text, date=None)` → id, `resolve_comment(id, *, resolved=True)`, `remove_comment(id)` | |
| `move_comment_to_text(id, anchor, *, occurrence=0)`, `move_comment(id, range)` | moves a thread onto exactly `anchor` (occurrences counted as `add_comment_on_text` does) or onto a `StoryRunRange` (a `RunRange` raises `TypeError`; a `Comment.anchor` read back works), keeping its id, author, date, text, replies and resolved flag; an emptied Google `goog_rdk` wrapper goes. A reply id, an unknown id or a text not found raises `RdocxError` |
| `accept_all()`, `reject_all()`, `accept_revision_id(id)`, `reject_revision_id(id)`, `accept_revisions_by_author(a)`, `reject_revisions_by_author(a)`, `accept_revisions_in_date_range(*, start, end)`, `reject_revisions_in_date_range(*, start, end)` | every supported story; dates as RFC 3339 strings |
| `compare(edited, author, timestamp, *, granularity="run", ignore_comments=False, ...)` → diagnostics (`ComparisonDiagnostic`: `location`, `message`) | turns `doc` into the redline of `doc` → `edited`; `granularity="word"` marks only the changed words, `ignore_comments=True` keeps `doc`'s comments and compares the rest |
| `rebuild_toc()` → `TocRebuildReport` (`entry_count`, `bookmark_count`, `diagnostics`, `diagnostic_count`) | the entry of a numbered heading gets a left stop after its number |
| `insert_toc(bi, max_level=3)` | inserts a TOC field (`TOC \o "1-N" \h`) at body index `bi`; `rebuild_toc()` then fills it with entries linked to the headings, their page numbers and the `TOC1`... styles |
| `update_page_fields()` → int, `update_layout_backed_fields()` → `LayoutBackedFieldUpdateReport` (`page_fields`, `num_pages_fields`, `page_reference_fields`, `section_fields`, `section_pages_fields`, `updated_count`, `diagnostics`, `diagnostic_count`), `update_fields(*, now=datetime, file_name=, file_path=, merge_fields=, ...)` → count, `update_fields_on_open` (get/set) | caches field results from rdocx's pagination; PAGE and NUMPAGES in headers and footers keep their cached results, as Word does on save (a diagnostic says so), and render with the right number anyway |
| `layout()` → tuple of `LayoutFragment` (`body_index`, `physical_page`, `displayed_page`, `bounds`: a `BoundingBox`, `x`, `y`, `width`, `height` in points), `layout_page(i)` → `LayoutPage` (`page_number`, `displayed_page_number`, `width`, `height`) | rdocx's pagination, on Word's line heights |
| `to_pdf(*, revision_view="accepted")` → bytes, `render_pages(*, dpi=150, format="png", quality=90, transparent=False, pages=None, revision_view="accepted")` → list of bytes, `render_page_to_png(i, dpi, *, revision_view="accepted")`, `render_all_pages(dpi, *, revision_view="accepted")`, `render_page_to_svg(i)` → `SvgRenderResult` (`svg` text, `diagnostics`: `SvgDiagnostic` `path`, `message`) or None past the last page, `to_pdfa_deterministic(profile="pdfa-2b")` → bytes (PDF/A-2b, or `"pdfa-3b"`) | `pages` zero-based. `revision_view="tracked"` shows tracked changes (deletions struck through, insertions underlined, a change bar); any other value raises `ValueError` |

## Paragraph, Run, Font, ParagraphFormat

- `Paragraph`: `text` (the accepted view: tracked insertions in, deletions out. Setting it leaves one run
  without direct formatting, keeps the paragraph style, format and comment ranges, and refuses a
  paragraph holding one end of a field that spans paragraphs, such as a TOC), `runs`, `style` (reads the
  style id. Set an id or a name the document defines, as python-docx does: `KeyError` when no paragraph
  style has it, `ValueError` for a character or table style), `alignment` (`WD_ALIGN_PARAGRAPH`),
  `numbering` (`(num_id, level)` or None, settable, unchecked), `paragraph_format`, `add_run(text)`,
  `add_hyperlink(text, url)` (appends at the end), `replace_text(old, new, *, expect=None)` → int (counted
  replacement in this paragraph only, body or table cell: across runs, comments kept; a count other than
  `expect` raises `ReplacementCountError` and changes nothing).
- `ParagraphFormat`: `space_before`, `space_after`, `line_spacing` (float for multiples, length for exact),
  `left_indent`, `right_indent`, `first_line_indent`, `keep_with_next`, `keep_together`,
  `page_break_before`, `widow_control`, `alignment`. None means inherited.
- `Run`: `text` (settable), `font`, `style_id`, `add_tab()` (a tab at the end of the run), `add_field(instruction,
  cached_result="")` (a simple field after the run, such as `"PAGE"`, showing `cached_result` until fields are
  updated).
- `Font`: `name`, `size` (length), `bold`, `italic`, `underline` (`WD_UNDERLINE` or bool), `strike`,
  `color` (set an `rdocx.RGBColor(r, g, b)`; reads back an `RGBColor`, a tuple: compare `str(font.color) ==
  "7B1E3A"`; a hex string is refused), `highlight` (Word colour name: "yellow", "green", ...), `shading`
  (hex fill). None means inherited.

## Tables

`Table`: `rows`, `cell(r, c)`, `clone_row(index, at=None)` (keeps cell and run formatting),
`remove_row(index)` (a comment inside the row goes with it), `style` (style id, unchecked), `alignment` (`WD_TABLE_ALIGNMENT`), `width` (settable).
`Row.cells` (a merged cell appears once: use `cells[-1]` for the last column); `Cell`: `text` (settable),
`replace_text(old, new, *, expect=None)` → int (as `Paragraph.replace_text`, over the cell), `paragraphs`, `add_paragraph(text)`, `width` (settable: writes the cell width, not the table grid),
`vertical_alignment` (`WD_CELL_VERTICAL_ALIGNMENT`), `shading` (hex fill, settable), `grid_span`,
`vertical_merge`, `border(edge)` and `set_border(edge, style, *, size, color)` (this cell), `margins` and
`set_margins(*, top, right, bottom, left)`. On the table: `set_cell_grid_span(r, c, n)` (consumes empty
cells only, and invalidates the table handle), `set_cell_vertical_merge(r, c, "restart" | "continue" |
None)`, `set_borders(style, *, size, color)`, `set_border(edge, style, *, size, color)`, `border(edge)` →
`(style, size, color)` or None, `set_cell_margins(*, top, right, bottom, left)`, `cell_margins` → `(top,
right, bottom, left)` or None, `set_column_width(c, w)` and `grid_widths` (the grid, get and set). Edges:
`top`, `bottom`, `left`, `right`, `insideH`, `insideV`; styles: `none`, `single`, `thick`, `double`,
`dotted`, `dashed`, `dotDash`, `wave`; `size` in eighths of a point, `color` hex. `Row`: `height`,
`height_rule` (`WD_ROW_HEIGHT_RULE.EXACTLY` or `AT_LEAST`), `is_header` (repeated at the top of each page),
`cant_split` (settable).

## Stories other than the body

`doc.story_items` lists the paragraphs, tables, pictures, content controls and fields of every story, with
`item.story.kind` in `body`, `table_cell`, `text_box`, `header`, `footer`, `footnote`, `endnote`, ... and
`item.story.part_name`. A header or footer has one story per variant part, and a text box as Word writes
it (`mc:AlternateContent`, DrawingML with a VML copy) is listed once. `doc.paragraphs` holds the body only,
and `rdocx text` and `text --json` (`stories`) print the others after it. `docx_ops.story_paragraphs(doc)`,
`all_text(doc)` and `count(doc, text)` read the package XML: every paragraph of every story, in the accepted
view, with its part and style id. A `field` item's `text` is the result the field shows, read across all its
runs for a complex field. A paragraph inside a content-control block is not listed (its control is), but the
item a `Comment.anchor` or `StoryRunPosition(paragraph=...).item` gives for it has its own `text` and `xml`.

## Read-only records

`Section` (`ordinal`, `is_final`, `orientation`, `page_width`, `page_height`, `margin_top`, `margin_right`,
`margin_bottom`, `margin_left`, `gutter`, `header_distance`, `footer_distance`, `column_count`,
`column_spacing`, `different_first_page`, `break_type`, `page_number_start`); `Style` (`style_id`, `name`,
`style_type`, `based_on`, `linked_style`, `next_style`, and the flags `is_default`, `priority`, `hidden`,
`semi_hidden`, `unhide_when_used`, `quick_format`, `locked`, `auto_redefine`); `HeaderFooterVariant` (`section_index`, `kind` header/footer, `variant`
default/first/even, `story`, `inherited`, `source_section`); `Hyperlink` (`url`, `anchor`, `text`,
`index_path`, `relationship_id`, `story`); `Comment` (`id`, `author`, `initials`, `date`, `text`,
`parent_id`, `resolved`, `anchor_text`: the accepted-view text the comment covers, through Google `goog_rdk`
wrappers, paragraphs joined with `"\n"`, `""` for a reference without a range, None for a reply or a
comment with no markers; `anchor`: its `StoryRunRange` or None); `Revision` (`id`, `kind`, `author`, `timestamp`, `story`); `StoryItem` (`kind` paragraph,
table, drawing, content_control, field, preserved_node; `index_path`, `direct_body_index`, `text`, `xml`,
`story`); `Story` (`kind`, `part_name`, `owner_index`).
