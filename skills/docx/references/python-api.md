# rdocx Python API (pinned build)

`import rdocx` with the pinned Python (`$R/python`). The surface looks like python-docx where it overlaps,
with document-level methods for everything python-docx lacks (comments, revisions, comparison, layout,
rendering). The modules have no docstrings: exact signatures are in the type stubs, `rdocx/_rdocx.pyi` next
to the installed module (`$R/python -c "import rdocx, os; print(os.path.dirname(rdocx.__file__))"`).

Lengths are EMU integers; build them with `rdocx.Pt(12)`, `rdocx.Inches(1)`, `rdocx.Cm(2)`, `rdocx.Mm(5)`,
`rdocx.Emu(n)` (each has `.pt`, `.inches`, `.cm`, `.mm`, `.emu`, `.twips`); widths, margins and font sizes
read back as `rdocx.Length`, the int subclass they share. Colours: every colour argument takes an
`rdocx.RGBColor(r, g, b)` (or `RGBColor.from_string("7B1E3A")`), a hex string with or without `#` (`"7B1E3A"`), or an `(r, g, b)` triple, and reads
back as an `RGBColor` (a tuple whose `str()` is the hex); a wrong type raises `TypeError`, a malformed value
`ValueError`. Errors: `rdocx.RdocxError`
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
`StoryRunRange` built on one) raise `StaleElementError` once an edit has retired them, and the message names
the call that did (`Document.remove_content invalidated it`). What each edit retires:

| Edit | Retires |
|---|---|
| formatting setters (font, paragraph format, style, alignment, numbering, widths, borders, shading), `Run.text` | nothing |
| appends and inserts: `add_paragraph`, `add_heading`, `add_table`, `add_page_break`, `insert_paragraph`, `paragraph.insert_paragraph_before`, `paragraph.add_run`, `table.add_row`, `table.add_column` | story items only: a held paragraph follows its paragraph, not its index |
| `paragraph.text = ...`, `cell.text = ...`, `replace_xml` on a paragraph, run, table or cell | the handles inside that element (the element's own handle stays valid), and story items |
| row and grid edits: `clone_row`, `remove_row`, `insert_column`, `remove_column`, `set_cell_grid_span`, `cell.split()` | that table's rows and cells (the `Table` stays valid), and story items |
| everything else: `remove_content`, `pop_content`, `move_content`, `clone_content`, `split_run`, `set_story_text`, `add_picture`, `add_comment`, `try_replace_text`, `replace_all_regex`, `paragraph.replace_text`, `cell.replace_text`, `replace_text_at` | every handle |

Re-navigate from `doc` after an edit of the last row (`doc.paragraphs[i].runs[j]`, `doc.story_items`); when
in doubt, re-fetch: a retired handle always raises, it never points at the wrong element.

## Document

| Member | Notes |
|---|---|
| `Document()`, `Document(path or stream)`, `Document.open(path)`, `Document.from_bytes(b)`, `doc.to_bytes()`, `doc.save(path or stream)` | `save("x.pdf")` (an extension of another format) raises `ValueError` naming `to_pdf`, and so does a table without rows. A new document has Word's usual styles (Normal, Title, Subtitle, Heading 1 to 9, No Spacing, Quote, List Paragraph, Caption, Table Grid). `save` writes a temporary file and renames it (`docx_ops.save_atomic` also refuses the input), and writes the package class its extension names: a .dotx saved as .docx becomes a document |
| `paragraphs`, `tables`, `sections`, `styles`, `stories`, `story_items`, `header_footer_variants`, `hyperlinks`, `comments`, `revisions` | read-only collections (sections and styles are snapshots, `styles["Normal"]` looks one up by name or id; `revisions` covers every story, `r.story.kind`) |
| `core_properties` → `CoreProperties` | `title`, `author`, `subject`, `keywords`, `category`, `comments`, `last_modified_by`, `content_status`, `language`, `identifier`, `version` (strings), `created`, `modified`, `last_printed` (`datetime` or None), `revision` (int): read and set |
| `add_style(name, style_type="paragraph", *, style_id=None, based_on=None, next_style=None, font_name=, font_size=, bold=, italic=, color=, space_before=, space_after=, left_indent=, right_indent=, first_line_indent=, underline=, alignment=, line_spacing=, keep_with_next=, keep_together=, page_break_before=, shading=, borders=, tab_stops=)` → `Style` | as python-docx's `styles.add_style`: unless given, the id keeps the letters, digits and hyphens of the name (`"Note box"` → `Notebox`). `based_on` and `next_style` take an id or a name. Lengths and `font_size` in EMU (`rdocx.Pt(11)`); `line_spacing` a float multiple or an exact length; `borders` maps an edge to `(style, eighths of a point, colour)`, `tab_stops` lists `(position, WD_TAB_ALIGNMENT[, WD_TAB_LEADER])`. `KeyError` for an unknown base, `ValueError` for a duplicate or for paragraph keywords on a character style |
| `set_style(style, *, based_on=, ...)` → `Style`, the keywords of `add_style` | changes an existing style, chosen by id or name; what it is not given keeps its value, and `borders` or `tab_stops` replace the whole set. It cannot remove a theme font or a theme colour the style already has, which Word keeps using over `font_name` / `color`. `KeyError` for an unknown style, and on any error the document is left unchanged |
| `remove_style(style)` → bool, `set_default_style(style)`, `default_font_name`, `default_font_size` (get and set) | the default font writes `w:docDefaults` (all four font slots, theme fonts there dropped); styles that set their own font, such as headings, keep it |
| `add_bullet_list_item(text, level=0)`, `add_numbered_list_item(text, level=0, *, restart=False)` → `Paragraph`, `restart_numbering(paragraph, start=1)` → `num_id` | append to the document's plain bullet or decimal list, continuing the last body paragraph on it (never a checklist or a numbered heading); `restart=True` and `restart_numbering` start a new count, as Word's "Restart at 1". `add_paragraph(text, style="List Bullet")` (or "List Number 2"...) works too: the style is created on demand |
| `add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.", start=1, left_indent=, hanging_indent=, font=), ...])` → definition id, `ListLevel.checklist(checked=False)`, `add_numbering_instance(definition_id, *, start=None, level=0)` → `num_id`, `link_style_to_numbering(style, num_id, level)` | a custom list: define its levels once, make an instance (`start` writes a `w:startOverride`), then `paragraph.numbering = (num_id, level)` or link a style to it. `checklist` gives a ballot-box level |
| `add_paragraph(text="", style=None)`, `add_heading(text, level=1)`, `add_page_break()`, `insert_paragraph(bi, text)`, `add_table(rows, cols, style=None)`, `insert_table(bi, rows, cols)` → `Table` | append / insert as python-docx does; heading level 0 is the Title style, 1 to 9 the Heading styles. Appends and inserts keep every held handle but story items |
| `add_picture(image, filename=None, width=None, height=None, *, after=None, description=, title=, decorative=False, name=, crop=, wrap="inline", position=, relative_to=)` → `StoryItem` | a picture in a new paragraph. `image`: bytes, a path or a binary stream; a width or a height alone keeps the aspect ratio. `description` is the alt text, `decorative=True` marks it decorative for screen readers, `crop=(left, top, right, bottom)` fractions, `wrap` `inline` / `square` / `tight` / `through` / `top_and_bottom` / `behind` / `in_front`, `position=(x, y)`, each a length or an alignment such as `"right"` or `"top"`, measured from `relative_to=(horizontal, vertical)` such as `("margin", "paragraph")` (the stub lists the values); options are checked before the image is embedded (`ValueError`) |
| `pictures` → tuple of `Picture` (`relationship_id`, `name`, `description`, `title`, `decorative`, `width`, `height`, `inline`, `content_type`, `filename`, `blob`), `set_picture_size(picture, width, height)` | body and table-cell pictures; a `Picture` resizes that picture only, a relationship id every picture showing it |
| `clone_content(handle, bi)` | copies a block to body index `bi`, with its formatting, fields and bookmarks (a copied bookmark is renamed, `MailMerge1`...); comment anchors are not copied |
| `move_content(source, destination)`, `pop_content(bi)` → `ContentFragment`, `insert_content(bi, fragment)`, `remove_content(bi)` | `destination` of `move_content` is counted before the move (moving the 2nd of ABCDE to 3 gives ACBDE). Removing content never leaves a comment without an anchor: a comment it covers whole goes with its replies; a removal that would cut a comment in part, and `pop_content` of content that carries a comment (fragments do not own threads), raise `RdocxError` and change nothing. Move the thread first, or use `move_content`, which keeps it on the moved block |
| `set_story_text(item, text)` | replaces a paragraph's text, keeping the first run's formatting (any story) |
| `split_run(body_index, run_index, character_offset)` | splits one run in two at a character offset; body paragraphs only (a table cell paragraph raises `ValueError`) |
| `update_section(i, *, margin_top=, orientation=, page_width=, ...)` → `Section`, `add_section(start_type=WD_SECTION.NEW_PAGE)` → `Section` | writes a section's page setup (`doc.sections` stays a snapshot: `section.left_margin = ...` raises `AttributeError` naming the keyword); lengths in EMU, `rdocx.Twips(1440)` for twips, and a nonzero page length under 0.01 inch raises `ValueError`. Switching the orientation swaps width and height. A new section copies the previous one's page size, margins and columns |
| `insert_section(i)`, `remove_section(i)` | `insert_section` adds an empty section at section index `i` (an empty paragraph carries its break: 0 before the first section, `len(doc.sections)` after the last one's content); `remove_section` removes that break, the paragraphs stay |
| `create_section_story(i, kind, variant)` → `Story`, `unlink_section_story(i, kind, variant)` → `Story`, `link_section_story(i, kind, variant, story)` → `Story` | `kind` `"header"` / `"footer"`, `variant` `"default"` / `"first"` / `"even"`. `create_section_story` gives section `i` a new, empty header or footer (`"first"` turns on `different_first_page`), which the next sections inherit (`doc.header_footer_variants`); fill it with `insert_content(story, fragment)` (a fragment from `pop_content`). `unlink_section_story` gives a section that inherits one its own copy, to edit apart; `link_section_story` makes it use `story` again |
| `try_replace_text(old, new, *, expect=None)` → int, `replace_all_regex([(pattern, repl), ...])` → int | literal and regex replacement across runs, in the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes (a Word text box once), headers and footers with their tables (once per variant part), footnotes and endnotes. A match across the edge of a content control, an insertion or a simple field is not replaced. With `expect=N`, a different count raises `ReplacementCountError` and changes nothing; `docx_ops.replace_batch` also refuses text left out of reach |
| `replace_all([(old, new), (old, new, expect), ...])` → tuple of counts | several literal replacements in one call, all or nothing: a pair whose `expect` is not met raises `ReplacementCountError` (naming the pair) and changes nothing |
| `replace_text_at(item, old, new, *, expect=None)` → int | the same replacement in one `StoryItem` only (`doc.story_items`): a body paragraph or table, a table cell's paragraph, a header, footer or footnote paragraph. A comment paragraph is its own text (not the anchor). A Word text box: both its copies (DrawingML and the VML fallback) are edited and counted once; copies that differ raise `RdocxError`, nothing changed. For one body paragraph or cell, `Paragraph.replace_text` and `Cell.replace_text` |
| `bookmarks` → tuple of `Bookmark` (`id`, `name`, `text`, `range`, `direct_range`, `issue`), `add_bookmark(name, range)` → id | `add_bookmark` takes a body `RunRange`; `direct_range` gives it back that way (None for a bookmark in a table cell or a content-control block), `range` counts paragraphs through tables; `issue` describes a broken marker, None otherwise |
| `set_header(text)`, `set_footer(text)`, `add_hyperlink_to_story(story, text, url)` | `set_header` / `set_footer` replace the default story's content (fields included); a comment anchored only in the replaced header or footer goes with it, replies included |
| `add_footnote(target, text)` → id, `add_endnote(target, text)` → id, `remove_footnote(id)`, `remove_endnote(id)`, `set_note_numbering(kind, *, number_format="decimal", start=1, restart="continuous", placement=None, section=None)` | `target` is a body paragraph or its last run: the reference ends it. Notes are written as Word writes them (reference styles, separators created when missing); `kind` `"footnote"` / `"endnote"`, `number_format` `decimal`, `upperRoman`, `lowerRoman`, `upperLetter`, `lowerLetter` |
| `page_color` (get and set), `set_text_watermark(text)`, `set_image_watermark(data, filename, width, height)`, `set_page_borders(section, style="single", *, width=, color=, space=, offset_from="page")` | page colour writes `w:background` and turns on its display; watermarks are VML shapes in the headers (Google Docs drops them) |
| `text()`, `to_markdown()`, `to_html()`, `to_odt()`, `to_rtf()`, `to_epub()` (bytes) | `text()` is what `rdocx text` prints (body, then the other stories), Markdown and HTML what `convert` writes. Content a format cannot hold emits an `rdocx.ConversionWarning` naming it |
| `word_count()`, `character_count(*, include_spaces=True)`, `page_count()` | body and tables, as Word counts them; pages from rdocx's layout |
| `validate()` → `ValidationReport` (`ok`, `errors`, `warnings`), `Document.validate_file(path)` → `ValidationReport` | the checks of `rdocx validate`, on the document as it would be saved now, or on a file that may not even open (a truncated part is an error naming it) |
| `render_template(data)` → tag count | `{{ path.to.value }}` tags (across runs, first run's format) and `{% for x in path %}` / `{% if path %}` blocks, each marker alone in its paragraph, closed by `{% endfor %}` / `{% endif %}`. Values render as JSON writes them: pass formatted strings. A missing path raises `RdocxError` naming the tag, nothing changed |
| `insert_document(other, at=None, *, conflict="reuse_equivalent")`, `copy_fragment(start, end=None)` → `DocumentFragment`, `import_fragment(fragment, at=None, *, conflict=)` | the body of another document (or body items `start` up to `end` excluded) with its styles, lists, pictures and links, before body index `at` (the end by default). `conflict="rename"` renames every style instead of reusing identical ones. Headers, footers and page setup of `other` are not copied |
| `custom_properties` → `CustomProperties`, `app_properties` → `AppProperties`, `settings` → `Settings` | custom properties: a dict-like mapping (`[k]`, `del`, `keys()`, `values()`, `items()`, `get(k)`) of str, int, float, bool or `datetime` (UTC); Google Docs drops them. App properties: `company`, `manager`, `template`, `application` (get and set), `pages`, `words`, `characters`, `characters_with_spaces`, `lines`, `paragraphs` (read-only). Settings: `track_revisions`, `odd_and_even_pages_header_footer` (get and set) |
| `content_controls` → tuple of `ContentControl` (`tag`, `alias`, `id`, `type`, `text`), `set_content_control_value(value, *, tag=None, alias=None)` → count | body controls; `type` such as `plain_text`, `rich_text`, `dropdown_list`, `combo_box`, `date`, `checkbox`. Setting a value does what Word does: the placeholder goes, a drop-down takes one of its items (else `ValueError` listing them), a date takes an ISO date (`w:fullDate`), a check box `"true"` / `"false"`; bound custom XML is written too. No match: `KeyError`; picture and group controls: `ValueError` |
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
  `insert_paragraph_before(text="", style=None)`, `replace_text(old, new, *, expect=None)` → int (counted
  replacement in this paragraph only, body or table cell: across runs, comments kept; a count other than
  `expect` raises `ReplacementCountError` and changes nothing), `xml` / `replace_xml(xml)` ("Raw XML").
- `add_hyperlink(text, url=None, *, anchor=None, tooltip=None)` appends a link: `url` for the web, `anchor`
  for a bookmark name or a heading `Paragraph` (which gets a `Heading_...` bookmark when it has none); a
  `url` starting with `#` names a bookmark too. An unknown bookmark raises `KeyError`. `Hyperlink.tooltip`
  reads it back.
- `add_equation(latex=None, *, mathml=None, display=False)` appends an equation from LaTeX or MathML
  (`ValueError` when it does not parse or would lose content); `equations` → `Equation` (`latex`,
  `mathml`, `display`, `diagnostics`). Body and table-cell paragraphs.
- `ParagraphFormat`: `space_before`, `space_after`, `line_spacing` (a bare number is a multiple, `2` double
  spacing as in python-docx; a length is exact), `left_indent`, `right_indent`, `first_line_indent`,
  `keep_with_next`, `keep_together`, `page_break_before`, `widow_control`, `alignment`, `outline_level` (0
  to 9), `right_to_left`, `shading` (fill), `set_border(edge, style="single", size=4, color=None)`,
  `border(edge)` → `(style, size, RGBColor)` or None, `remove_border(edge)`, `clear_borders()` (edges `top`,
  `bottom`, `left`, `right`, `between`, `bar`), `tab_stops` (as python-docx: `add_tab_stop(position,
  alignment=WD_TAB_ALIGNMENT.LEFT, leader=WD_TAB_LEADER.SPACES)` keeps position order, a second stop at one
  position raises `ValueError`; `tab_stops[i]` is a live `TabStop` with settable `position`, `alignment`,
  `leader`; `del tab_stops[i]`, `clear_all()`, `len`). None means inherited.
- `Run`: `text` (settable), `font`, `bold`, `italic`, `underline` (shortcuts for `font.*`), `style_id`,
  `add_tab()` (a tab at the end of the run), `add_break(WD_BREAK.LINE)` (`PAGE`, `COLUMN`, `TEXT_WRAPPING`...;
  refused in headers and footers), `add_field(instruction, cached_result="")` (a simple field after the
  run, such as `"PAGE"`, showing `cached_result` until fields are updated), `add_picture(image, width=None,
  height=None, *, filename=, description=, ...)` → `Picture` (the options of `Document.add_picture`; body
  and table-cell runs: in a header or footer it raises, gap header-footer-picture), `xml` / `replace_xml(xml)`.
- `Font`: `name`, `size` (length), `bold`, `italic`, `underline` (`WD_UNDERLINE` or bool), `strike`,
  `double_strike`, `all_caps`, `small_caps` (written, not drawn by rdocx's renderer), `hidden`,
  `superscript` / `subscript` (python-docx's tri-state: setting one False clears only itself),
  `character_spacing` (a length, negative condenses), `highlight_color` (`WD_COLOR_INDEX.YELLOW`...) or
  `highlight` (Word colour name: "yellow", "green", ...), `language`, `east_asian_language`,
  `complex_script_language` (tags such as `"fr-FR"`, checked), `east_asian_name`, `complex_script_name`
  (fonts that `name` also sets), `rtl`, `color` (a `ColorFormat` as in python-docx: `rgb`, `type`
  `MSO_COLOR_TYPE` RGB / THEME / AUTO, `theme_color` `MSO_THEME_COLOR`, get and set; `font.color = ...` is
  a shortcut for `font.color.rgb = ...`, `"auto"` and None accepted), `shading` (fill, `RGBColor` or None).
  None means inherited.

## Tables

`Table`: `rows`, `columns`, `cell(r, c)`, `add_row()` → `Row` (its `cells` can be written one after the
other), `add_column(width=None)` → `Column` (`index`, `width`, `cells`), `insert_column(i, width=None)`,
`remove_column(i)`, `clone_row(index, at=None)` (keeps cell and run formatting), `remove_row(index)` (a
comment inside the row goes with it), `style` (an id or a name; a name the document lacks is still written,
with a `UserWarning` that Word will draw the default style), `alignment` (`WD_TABLE_ALIGNMENT`), `width`
(settable), `indent` (a length, negative pulls the table left, None removes it), `autofit` (False writes a
fixed layout), `first_row`, `last_row`, `first_col`, `last_col`, `horz_banding`, `vert_banding` (the style's
conditional formats, python-pptx names), `xml` / `replace_xml(xml)`.
`Row.cells` (a merged cell appears once: use `cells[-1]` for the last column); `Cell`: `text` (settable),
`replace_text(old, new, *, expect=None)` → int (as `Paragraph.replace_text`, over the cell), `paragraphs`,
`add_paragraph(text)`, `add_table(rows, cols)` (a nested table sized to the cell) and `tables`, `split()`
(undoes a horizontal merge, returns the cell count), `width` (settable: writes the cell width, not the
table grid), `vertical_alignment` (`WD_CELL_VERTICAL_ALIGNMENT`), `shading` (fill, settable, reads an
`RGBColor`), `grid_span`, `vertical_merge`, `border(edge)` and `set_border(edge, style, *, size, color)`
(this cell), `margins` and `set_margins(*, top, right, bottom, left)`, `xml` / `replace_xml(xml)`. On the
table: `set_cell_grid_span(r, c, n)` (consumes empty cells only), `set_cell_vertical_merge(r, c, "restart" |
"continue" | None)`, `set_borders(style, *, size, color)`, `set_border(edge, style, *, size, color)`,
`border(edge)` → `(style, size, RGBColor)` or None, `set_cell_margins(*, top, right, bottom, left)`,
`cell_margins` → `(top, right, bottom, left)` or None, `set_column_width(c, w)` and `grid_widths` (the grid,
get and set). Edges: `top`, `bottom`, `left`, `right`, `insideH`, `insideV`; styles: `none`, `single`,
`thick`, `double`, `dotted`, `dashed`, `dotDash`, `wave`; `size` in eighths of a point. `Row`: `height`,
`height_rule` (`WD_ROW_HEIGHT_RULE.EXACTLY` or `AT_LEAST`), `is_header` (repeated at the top of each page),
`cant_split` (settable). In a nested table, `paragraphs`, `add_paragraph`, `replace_text`, `clone_row`,
`remove_row`, `remove_column` and raw XML raise `NotImplementedError`; `cell.text`, formatting and grid
edits work.

## Headers and footers, the python-docx way

`doc.sections[i].header`, `footer`, `first_page_header`, `first_page_footer`, `even_page_header`,
`even_page_footer` are live `HeaderFooter` handles (`kind`, `variant`, `section_index`), valid across edits:

| Member | Notes |
|---|---|
| `paragraphs`, `add_paragraph(text="", style=None)` | a new header holds one empty paragraph, as in python-docx: write `header.paragraphs[0].text`. pass `style="Footer"` (or `"Header"`) for Word's look: a new paragraph gets no style by itself |
| `add_page_number(template="Page {PAGE} of {NUMPAGES}", *, alignment=CENTER)` → `Paragraph` | `{PAGE}`, `{NUMPAGES}`, `{SECTIONPAGES}` fields that Word, Google Docs and rdocx fill on every page |
| `add_table(rows, cols, width=None)` → `HeaderFooterTable`, `tables` | `cell(r, c)` and `rows[i].cells` (`text`, `paragraphs`, `add_paragraph`), `row_count`, `column_count`, `style` |
| `is_linked_to_previous` (get and set) | setting False gives the section an empty story of its own; writing into a linked one edits the earlier section's story |

Writing into a first-page story turns `different_first_page_header_footer` on for the section, into an even
story turns `doc.settings.odd_and_even_pages_header_footer` on (Word ignores even headers without it). No
picture in a header or footer from Python yet (gap header-footer-picture), and raw XML is refused there.
The lower-level story calls (`create_section_story`, `insert_content`) stay available.

## Raw XML

When a feature has no API, edit the element's XML instead of reaching for lxml: `Paragraph`, `Run`, `Table`
and `Cell` have `xml` (bytes of the standalone element) and `replace_xml(xml)` (str or bytes, one element
of the same kind), and `doc.section_xml(i)` / `doc.replace_section_xml(i, xml)` do the same for a section's
`w:sectPr`. Prefixes the document declares need no declaration. A replacement is refused with
`ValueError`, the document unchanged, when the XML is malformed or has a DOCTYPE, holds another or a second
element, puts an element where Word refuses it, uses an undeclared foreign namespace, leaves a cell without
a paragraph, references a relationship id the part lacks or of the wrong type, or adds or removes a section
break. Body and table-cell elements only: in a header, a footer or a nested table it raises
`NotImplementedError`. The element's own handle stays valid; the handles inside it retire.

## Values that would do nothing raise

A bare integer is read as EMU, so `font.size = 12`, `space_after = 6` or a margin of `1440` would be
invisible: they raise `ValueError` naming `Pt` (or `Twips` / `Inches` for page lengths). `section.left_margin
= ...` raises `AttributeError` naming the `update_section` keyword, and a python-docx private such as
`paragraph._p` raises `AttributeError` pointing at `xml` / `replace_xml`.

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
`column_spacing`, `different_first_page`, `break_type`, `page_number_start`; and the live `header` ... handles
above); `Style` (`style_id`, `name`,
`style_type`, `based_on`, `linked_style`, `next_style`, and the flags `is_default`, `priority`, `hidden`,
`semi_hidden`, `unhide_when_used`, `quick_format`, `locked`, `auto_redefine`); `HeaderFooterVariant` (`section_index`, `kind` header/footer, `variant`
default/first/even, `story`, `inherited`, `source_section`); `Hyperlink` (`url`, `anchor`, `text`, `tooltip`,
`index_path`, `relationship_id`, `story`); `Comment` (`id`, `author`, `initials`, `date`, `text`,
`parent_id`, `resolved`, `anchor_text`: the accepted-view text the comment covers, through Google `goog_rdk`
wrappers, paragraphs joined with `"\n"`, `""` for a reference without a range, None for a reply or a
comment with no markers; `anchor`: its `StoryRunRange` or None); `Revision` (`id`, `kind`, `author`, `timestamp`, `story`); `StoryItem` (`kind` paragraph,
table, drawing, content_control, field, preserved_node; `index_path`, `direct_body_index`, `text`, `xml`,
`story`); `Story` (`kind`, `part_name`, `owner_index`).
