# rdocx Python API (pinned build)

`import rdocx` with the pinned Python (`$R/python`). The surface borrows python-docx's shapes (`paragraphs`,
`runs`, `font`, `paragraph_format`, `tables`) but not all its names (next section), and has document-level
methods for everything python-docx lacks (comments, revisions, comparison, layout, rendering). Exact signatures
are in the type stubs, `rdocx/_rdocx.pyi` next to the installed module (`$R/python -c "import rdocx, os;
print(os.path.dirname(rdocx.__file__))"`).

**Lengths are EMU integers: build every one with `rdocx.Pt(12)`, `rdocx.Inches(1)`, `rdocx.Cm(2)`,
`rdocx.Mm(5)`, `rdocx.Twips(1440)` or `rdocx.Emu(n)`.** A bare int is taken as EMU: `font.size = 12` raises
`ValueError`, but `space_after = 6`, `character_spacing = 2` and a margin of 1440 silently save as about 0, and
`line_spacing = 2` an exact spacing of 0 (gap docx-bare-int-length-saved-as-zero); a line spacing multiple is a float
(`2.0`). Each length has `.pt`,
`.inches`, `.cm`, `.mm`, `.emu`, `.twips`, and widths, margins and sizes read back as `rdocx.Length`, the int
subclass they share. Colours: an `rdocx.RGBColor(r, g, b)` (or `RGBColor.from_string("7B1E3A")`), a hex string
with or without `#`, or an `(r, g, b)` triple; they read back as an `RGBColor` (a tuple whose `str()` is the
hex); a wrong type raises `TypeError`, a malformed value `ValueError`. Errors: `rdocx.RdocxError` (base),
`XmlError`, `PackageError`, `LayoutError`, `StaleElementError`, `ReplacementCountError` (an `expect` count not
met); lookups by text raise `ValueError`, refused options `ValueError` or `RdocxError`.

## python-docx names

These python-docx calls work as they do there: `doc.add_heading(text, level)` (level 0 is Title), `add_paragraph(text="",
style=None)` (a style the document has: a new one lacks List Bullet and List Number, below), `add_table(rows, cols, style=None)`, `doc.add_page_break()`, `p.insert_paragraph_before(text, style=None)`
(an insert: it retires every handle), `run.bold` / `italic` / `underline`, `run.add_break(WD_BREAK.PAGE)`,
`font.highlight_color = WD_COLOR_INDEX.YELLOW` (`rdocx.enum.text`), `WD_SECTION` and `WD_ORIENT` (alias
`WD_ORIENTATION`, `rdocx.enum.section`), and a binary stream for `Document(f)`, `save(f)` and `add_picture(f)`. These
need a rewrite; the ones marked "raises" fail with an error naming the rdocx way, the others with a bare
`AttributeError` or `TypeError` (gap python-docx-names):

| python-docx | rdocx |
|---|---|
| `add_paragraph(text, style="List Bullet")`, `style="List Number"` (`KeyError` on a new document) | `doc.add_bullet_list_item(text)`, `doc.add_numbered_list_item(text)`, or a template that has those styles |
| `table.add_column(w)`, `table.columns` (raise) | `t.insert_column(i, before=False)` then `t.set_column_width(i + 1, w)`; `len(t.rows[0].cells)` |
| `font.color.type == MSO_COLOR_TYPE.RGB` | `font.color.type == "rgb"` (`"theme"`...) |
| `doc.sections[i].header` (raises), `section.left_margin = x`, `doc.add_section(WD_SECTION.NEW_PAGE)` (raises) | `doc.section(i).header`, `doc.section(i).left_margin = x` (or `update_section`), `doc.insert_section(i)` |
| a new header's `paragraphs[0]` | `doc.section(i).header.add_paragraph(text)`: a new header holds no paragraph |
| `doc.styles["Normal"]` (raises), and changing it | `next(s for s in doc.styles if s.name == "Normal")` (a snapshot); `doc.set_style("Normal", font_size=rdocx.Pt(11))` |
| `tab_stops[i].alignment = ...` | `del tab_stops[i]`, then `add_tab_stop(position, alignment, leader)` |
| `doc.custom_properties["k"] = v` | `custom_properties.set("k", v)`, `.get("k")`, `.delete("k")` |
| `add_picture` with a width alone | both sizes, `add_picture(image, filename, w, h)`, or neither (the image's own size) |
| `paragraph._p`, `run._r` | `paragraph.xml` and `replace_xml(xml)` ("Raw XML") |

## Three index spaces

| Space | Used by | Counts |
|---|---|---|
| body index | `find_content_index(handle or text)`, `find_content_indices(text)`, `insert_paragraph`, `insert_table`, `remove_content`, `pop_content`, `insert_content`, `clone_content` and `move_content` destinations, `insert_document(at=)`, `split_run(i, ...)`, `RunPosition.body_index`, `LayoutFragment.body_index`, `StoryItem.direct_body_index`, `rdocx text --json` `body_index`, `rdocx comment add --start-paragraph` | top-level children of the body: paragraphs, tables, content-control blocks (the TOC is one block) |
| flow index | `doc.paragraphs[i]` | every paragraph outside tables, those inside content-control blocks (TOC entries) included, and the cell paragraphs of a table inside such a block |
| story path | `doc.story_items`, `StoryItem.index_path`, `Hyperlink.index_path`, `set_story_text(item, ...)`, `ContentLocation(story, kind, index_path)` | per story (body, table cell, text box, header, footer, footnote, ...): like the body index plus items for pictures, content controls and fields |

- `find_content_index(handle)` raises `ValueError` for a paragraph that is not a direct child of the body
  (a TOC entry, a paragraph in a content-control block or a table cell).
- `find_content_indices(text)` returns every top-level block whose text contains `text`, the TOC block
  included for any heading (never delete by it); it does not search tables. `find_content_index(text)`
  returns the first, or raises `ValueError`.
- Run indices of `RunPosition` and `split_run` count the runs `Paragraph.runs` lists, those inside an inline
  content control or a tracked insertion included. `Paragraph.runs` does not hold the text of a simple
  field, a smart tag or a custom XML element, which `Paragraph.text` holds: a character offset in `p.text`
  is not a run offset. `docx_ops.locate` returns run offsets, and `docx_ops.comment_on_text` finds a text and
  anchors on it.

## Handles

Every handle is checked: `Paragraph`, `Run`, `Table`, `Row`, `Cell`, `StoryItem`, `doc.section(i)` and the
`HeaderFooter` and `FootnoteEndnote` reached from them raise `StaleElementError` once an edit has retired
them, and the message names the call that did (`First invalidated by Document.add_paragraph`). What each edit
retires:

| Edit | Retires |
|---|---|
| formatting setters (font, paragraph format, style, alignment, numbering, cell width and shading, table style, borders), `run.add_tab()`, `run.add_break()`, `add_bookmark`, `add_style`, `set_style`, page setup (`update_section`, `doc.section(i).margin_top = ...`), properties, saving, rendering, reading | nothing |
| `Run.text = ...` | story items |
| `replace_xml` on a paragraph, run, table or cell | the handles inside that element (its own handle stays valid) |
| appends: `add_paragraph`, `add_table`, `paragraph.add_run`, `add_hyperlink`, `add_heading_hyperlink`, `paragraph.text = ...`, `cell.text = ...`, `add_footnote`, `add_endnote`, `add_chart`, a header's or footer's `add_paragraph` (which also retires the `HeaderFooter` handles) | every `Run` handle and story items; paragraphs, tables, rows, cells and `doc.section(i)` stay valid |
| `table.add_row()`, `table.remove_row(i)` | the same, and that table's rows and cells (the `Table` stays valid) |
| everything else: `insert_paragraph`, `insert_table`, `add_bullet_list_item`, `add_numbered_list_item`, `run.add_field`, `run.remove()`, `cell.add_paragraph`, `clone_row`, `insert_column`, `delete_column`, `set_cell_grid_span`, `cell.merge`, every replacement (`try_replace_text`, `replace_all`, `replace_all_regex`, `replace_text`, `replace_text_at`), comments, `remove_content`, `pop_content`, `move_content`, `clone_content`, `split_run`, `set_story_text`, `add_picture`, `insert_html`, `set_header`, `set_footer`, accepting or rejecting revisions, the font table (`embed_font`, `set_font`, `remove_font`, `remove_embedded_font`, the embedding flags) | every handle |

Re-navigate from `doc` after an edit (`doc.paragraphs[i].runs[j]`, `doc.tables[t].cell(r, c)`,
`doc.section(i).header`, `doc.story_items`): `doc.paragraphs[2].runs[0].font.bold = True` always works, a run
held across an append does not. When in doubt, re-fetch: a retired handle always raises, it never points at
the wrong element. A note's handle cannot be fetched again (gap docx-note-reference-and-removal): finish with
it before the next edit.

Collect a table's cells once, then edit them: `cell.text = ...` keeps rows and cells valid, and walking
`table.rows` and `row.cells` costs time that grows with the square of the document on this build (400 cells in
a short document: 3 s; 1,600 cells: 50 s; tensorbee/rdocx#378). Re-walk only after an edit that retires them.

## Document

| Member | Notes |
|---|---|
| `Document()`, `Document(path or stream)`, `Document.open(path or stream)`, `Document.from_bytes(b)`, `doc.to_bytes()`, `doc.save(path or stream)` | a stream is a binary file object (`io.BytesIO`, `open(f, "rb")`). A new document has Word's usual styles (Normal, Title, Subtitle, Heading 1 to 9, No Spacing, Quote, List Paragraph, Caption, Table Grid). `save` writes the package class its extension names (a .dotx saved as .docx becomes a document), and any other extension too: `save("x.pdf")` writes a DOCX under that name (gap docx-invalid-values-accepted), so convert with `to_pdf()`. `docx_ops.save_atomic(doc, out, src)` also refuses the input |
| `Document.from_markdown(text, *, template=None, images=())`, `Document.open_markdown(path, *, template=None, images=())` → `(Document, diagnostics)` | CommonMark with GitHub tables and strikethrough: headings get Heading N, lists numbering, tables a header row, links hyperlinks, images inline pictures scaled to the text column, block quotes Quote, code HTML Preformatted. Images come from `images=[(name, data, content_type), ...]` only (`[("pic.png", data, "image/png")]` for `![x](pic.png)`), not from the disk: `rdocx convert` reads the files next to the input. `template` is a `Document` or a path whose styles, numbering, headers and footers the new one takes. Each `HtmlDiagnostic` (`location` such as `"bytes 0..23"`, `message`) names what was dropped or approximated |
| `Document.from_html(html, *, images=())`, `Document.open_html(path, *, images=())` → `(Document, diagnostics)`, `doc.insert_html(html, *, after=paragraph, images=())` → diagnostics | HTML the same way: safe links and data-URI images kept (a `javascript:` link loses its target), `insert_html` after a body `Paragraph` (required: `doc.paragraphs[-1]` for the end). Each `HtmlDiagnostic` has `location` (`html/body/p[1]/a[2]`), `message` and the CSS `property` it is about, or None. Both importers keep a space before or after an underlined, struck, bold or italic span outside it, collapse spaces on both sides of a span edge into one (`a <s>b </s> c` reads `a b c`), and keep `&nbsp;` |
| `paragraphs`, `tables`, `sections`, `styles`, `stories`, `story_items`, `header_footer_variants`, `hyperlinks`, `comments`, `revisions`, `bookmarks`, `fonts` | read-only collections (sections and styles are tuples of snapshots; `revisions` covers every story, `r.story.kind`) |
| `pictures()`, `content_controls()` | methods, each a tuple (below) |
| `core_properties` → `CoreProperties` | `title`, `author`, `subject`, `keywords`, `category`, `comments`, `last_modified_by`, `content_status`, `language`, `identifier`, `version` (strings), `created`, `modified`, `last_printed` (`datetime` or None), `revision` (int): read and set |
| `add_style(name, style_type="paragraph", *, style_id=None, based_on=None, next_style=None, font_name=, font_size=, bold=, italic=, color=, space_before=, space_after=, left_indent=, right_indent=, first_line_indent=, underline=, alignment=, line_spacing=, keep_with_next=, keep_together=, page_break_before=, shading=, borders=, tab_stops=)` → `Style` | as python-docx's `styles.add_style`: unless given, the id keeps the letters, digits and hyphens of the name (`"Note box"` → `Notebox`). `based_on` and `next_style` take an id or a name. Lengths and `font_size` in EMU (`rdocx.Pt(11)`); `line_spacing` a float multiple or a length; `borders` maps an edge to `(style, eighths of a point, colour)`, `tab_stops` lists `(position, WD_TAB_ALIGNMENT, WD_TAB_LEADER or None)`. `KeyError` for an unknown base, `ValueError` for a duplicate or for paragraph keywords on a character style |
| `set_style(style, *, based_on=, ...)` → `Style`, the keywords of `add_style` | changes an existing style, chosen by id or name; what it is not given keeps its value; `borders` sets the edges it names (the others stay), `tab_stops` replaces the whole set. It cannot remove a theme font or a theme colour the style already has, which Word keeps using over `font_name` / `color`. `KeyError` for an unknown style, and on any error the document is left unchanged |
| `remove_style(style)` → bool, `set_default_style(style)`, `default_font_name`, `default_font_size` (get and set) | the default font writes `w:docDefaults`; styles that set their own font, such as headings, keep it |
| `add_bullet_list_item(text, level=0)`, `add_numbered_list_item(text, level=0)` → `Paragraph` | append to the document's plain bullet or decimal list, continuing the last body paragraph on it. No restart (gap docx-list-numbering-options): a new count is a new instance, below |
| `add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.", start=1, left_indent=, hanging_indent=), ...])` → definition id, `ListLevel.checklist(checked=False)`, `add_numbering_instance(definition_id, *, start=None, level=0)` → `num_id`, `link_style_to_numbering(style, num_id, level)` | a custom list: define its levels once, make an instance (`start` writes a `w:startOverride`), then `paragraph.numbering = (num_id, level)` or link a style to it. `checklist` gives a ballot-box level |
| `add_paragraph(text="", style=None)`, `insert_paragraph(bi, text, style=None)`, `add_heading(text="", level=1)`, `add_page_break()` → `Paragraph`, `add_table(rows, cols, style=None)`, `insert_table(bi, rows, cols)` → `Table` | append or insert at body index `bi`; `style` an id or a name (`KeyError` when the document lacks it); `add_heading` uses Title for level 0, Heading N otherwise |
| `add_picture(image, filename=None, width=None, height=None, *, after=None, options=None)` → `Picture` | a picture in a new paragraph at the end of the body, or after a `StoryItem` or `Picture`. `image`: bytes, a path or a binary stream. Both sizes or neither (the image's own size); one alone raises `RdocxError` (gap python-docx-names). `options=rdocx.PictureOptions(description=, title=, decorative=False, name=, crop=(left, top, right, bottom), placement="inline", horizontal_relative_from="page", vertical_relative_from="page", horizontal_offset=, vertical_offset=, horizontal_alignment=, vertical_alignment=, distances=, relative_height=)`: `description` is the alt text, `decorative=True` marks it decorative, `crop` in fractions, `placement` `inline` / `square` / `tight` / `top_and_bottom` / `behind` / `in_front`, alignments such as `"right"` or `"top"`, offsets in EMU, `relative_from` such as `"margin"` or `"paragraph"` (`ValueError` for an unknown one) |
| `pictures()` → tuple of `Picture` (`relationship_id`, `name`, `description`, `title`, `decorative`, `anchored`, `width`, `height`, `content_type`, `bytes`, `media_part_name`, `external_url`, `story`, `occurrence`, `location`), `set_picture_size(picture, width, height)`, `set_picture_sizes_for_relationship(rid, width, height)` → count | every picture of the body, table cells, headers and footers (`story.kind`); a `Picture` resizes that picture only, a relationship id every picture showing it |
| `clone_content(handle, bi)` | copies a block to body index `bi`, with its formatting, fields and bookmarks (a copied bookmark is renamed, `MailMerge1`...); comment anchors are not copied |
| `move_content(source, destination)`, `pop_content(bi)` → `ContentFragment`, `insert_content(bi, fragment)`, `remove_content(bi)` | `destination` of `move_content` is counted before the move (moving the 2nd of ABCDE to 3 gives ACBDE). Removing content never leaves a comment without an anchor: a comment it covers whole goes with its replies; a removal that would cut a comment in part, and `pop_content` of content that carries a comment, raise `RdocxError` and change nothing. Move the thread first, or use `move_content`, which keeps it on the moved block |
| `set_story_text(item, text)` | replaces a paragraph's text, keeping the first run's formatting (any story); a `\t` or `\n` in `text` is written as a tab or a line break |
| `split_run(body_index or paragraph, run_index, character_offset)` | splits one run in two at a character offset; body paragraphs only |
| `update_section(i, *, margin_top=, orientation=, page_width=, ...)` → `Section`, `section(i)` → `SectionHandle`, live | `doc.sections` are snapshots; write the page setup with `update_section` or on `doc.section(i)` (`margin_top`, `left_margin`, `page_width`, `orientation = "landscape"`, `gutter`, `header_distance`, `different_first_page_header_footer`...). Switching the orientation swaps width and height |
| `insert_section(i)`, `remove_section(i)` | `insert_section` adds a section break at section index `i` (an empty paragraph carries it: `len(doc.sections)` after the last one's content); a new section copies the previous one's page size, margins and columns; `remove_section` removes that break, the paragraphs stay |
| `create_section_story(i, kind, variant)` → `Story`, `unlink_section_story(i, kind, variant)` → `Story`, `link_section_story(i, kind, variant, story)` → `Story` | the lower-level header and footer calls: `kind` `"header"` / `"footer"`, `variant` `"default"` / `"first"` / `"even"`; fill a story with `insert_content(story, fragment)` |
| `try_replace_text(old, new, *, expect=None)` → int, `replace_all_regex([(pattern, repl), ...])` → int, `replace_all_regex_expected([(pattern, repl, expect), ...])` → counts | literal and regex replacement across runs, in the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes (a Word text box once), headers and footers with their tables (once per variant part), footnotes and endnotes. A match across the edge of a content control, an insertion or a simple field is not replaced. With `expect=N`, a different count raises `ReplacementCountError` and changes nothing; `docx_ops.replace_batch` also refuses text left out of reach |
| `replace_all([(old, new), (old, new, expect), ...])` → tuple of counts | several literal replacements in one call, all or nothing: a pair whose `expect` is not met raises `ReplacementCountError` (naming the pair) and changes nothing |
| `replace_text_at(item, old, new, *, expect=None)` → int | the same replacement in one `StoryItem` only (`doc.story_items`): a body paragraph or table, a table cell's paragraph, a header, footer or footnote paragraph. A Word text box: both its copies (DrawingML and the VML fallback) are edited and counted once; copies that differ raise `RdocxError`, nothing changed. For one body paragraph or cell, `Paragraph.replace_text` and `Cell.replace_text` |
| `bookmarks` → tuple of `Bookmark` (`id`, `name`, `text`, `range`, `direct_range`, `issue`), `add_bookmark(name, range)` → id | `add_bookmark` takes a body `RunRange`; `direct_range` gives it back that way (None for a bookmark in a table cell or a content-control block), `range` counts paragraphs through tables; `issue` describes a broken marker, None otherwise |
| `set_header(text)`, `set_footer(text)`, `add_hyperlink_to_story(story, text, url)` | `set_header` / `set_footer` replace the default story's content (fields included); a comment anchored only in the replaced header or footer goes with it, replies included |
| `add_footnote(reference, text="")`, `add_endnote(reference, text="")` → `FootnoteEndnote` (`id`, `paragraphs`, `tables`, `add_paragraph(text)`, `add_table(r, c)`, `delete()`), `set_note_policy(family, *, format="decimal", start=None, restart="continuous", placement=None)` | `reference` is a body paragraph or its last run: the reference ends it. `family` `"footnote"` / `"endnote"`, `format` `decimal`, `upperRoman`, `lowerRoman`, `upperLetter`, `lowerLetter`; `doc.section(i).set_note_policy` for one section. The reference is written unstyled, and the handle is the only way to delete a note (gap docx-note-reference-and-removal) |
| `set_page_color("FDF6E3")`, `set_text_watermark(text)`, `set_image_watermark(data, filename, width, height)`, `doc.section(i).set_page_border(edge, *, style="single", width=8, space=0, color="000000")`, `doc.section(i).remove_watermark()` | page colour: six hex digits, no `#`, writes `w:background` and turns on its display (no getter); watermarks are VML shapes in the headers (Google Docs drops them), `doc.section(i).set_text_watermark(text, width=, height=, font_size=, rotation=, color=, opacity=)` for one section; a page border per edge `top`, `left`, `bottom`, `right`, `width` in eighths of a point |
| `text()`, `text_report()` → `TextReport` (`text`, `scope`, `diagnostics`: `TextDiagnostic`), `to_markdown()`, `to_html()` (str), `to_odt()` → `OdtExportResult`, `to_rtf()` → `RtfExportResult`, `to_epub()` → `EpubExportResult` (each `bytes`, `diagnostics`) | `text()` is what `rdocx text` prints (body, then the other stories), Markdown and HTML what `convert` writes. What a format cannot hold is a diagnostic of the result (`OdtDiagnostic` and `EpubDiagnostic`: `path`, `message`; `RtfDiagnostic`: `offset`, `destination`, `message`): read them |
| `word_count()`, `character_count(*, include_spaces=True)`, `page_count(*, font_dir=None)` | body and tables, as Word counts them; pages from rdocx's layout |
| `validate()` → `ValidationReport` (`is_valid`, `findings`: `ValidationFinding` `severity` `"error"` / `"warning"`, `message`), `rdocx.validate_file(path)`, `rdocx.validate_bytes(data)` | the checks of `rdocx validate`, on the document as it would be saved now, or on a file that may not even open (a truncated part is an error naming it). A Word table of contents makes it raise (gap docx-validate-paraid-in-content-control) |
| `sign(private_key, certificate)` → report, `verify_signatures()` → tuple of reports | a package signature over every part and relationship, written by the next `save`: `private_key` and `certificate` are the bytes of a PKCS#8 RSA key and an X.509 certificate (DER or one unencrypted PEM block each). A report is a dict: `signature_part`, `signer` (`subject`, `issuer`, `serial_number`), `cryptographically_valid`, `coverage_complete`, `covered_parts`, `covered_relationships`, `issues` (`kind` such as `changed_reference`, `details`); `verify_signatures()` checks the document as it is now (an empty tuple when unsigned), so any edit after signing shows. A signed package or one with external relationships raises `PackageError`. Certificate trust is not checked (`references/cli.md`, Signatures) |
| `render_template(data)` → tag count | `{{ path.to.value }}` tags (across runs, first run's format) and `{% for x in path %}` / `{% if path %}` blocks, each marker alone in its paragraph, closed by `{% endfor %}` / `{% endif %}`. Values render as JSON writes them: pass formatted strings. A missing path raises `RdocxError` naming the tag, nothing changed |
| `insert_document(other, *, at, policy=None)`, `import_fragment(destination, fragment, *, policy=None)`, `DocumentFragment.from_range(document, start, end)` | the body of another document before body index `at` (`len(doc.paragraphs) + len(doc.tables)` for the end of a plain body), with its styles, lists, pictures and links; or the body items of `other` from `start` up to `end` excluded, both `ContentLocation(story, "paragraph", [i])` on `other`'s body story (`next(s for s in other.stories if s.kind == "body")`), imported at a `ContentLocation` of `doc` (`ContentLocation.end(story)` for the end). `policy=rdocx.FragmentConflictPolicy.rename_all()` renames every style instead of reusing identical ones (`reuse_equivalent()`, the default; `.with_style_reuse(False)`...). Headers, footers and page setup of `other` are not copied |
| `custom_properties` → `CustomProperties` (`get(name)` → `CustomProperty` or None, `set(name, value)`, `delete(name)` → bool), `app_properties` → `AppProperties`, `settings` → `Settings` | custom properties: `value` str, int, float, bool or `rdocx.CustomFileTime("2026-10-31T00:00:00Z")` for a date; `CustomProperty` has `name`, `type` (`lpwstr`, `i4`, `r8`, `bool`, `filetime`...) and `value`; no listing in Python (`rdocx meta get --json` lists them), no `[k]` (gap python-docx-names); Google Docs drops them. App properties: `company`, `manager`, `template`, `application` (get and set), `pages`, `words`, `characters`, `characters_with_spaces`, `lines`, `paragraphs` (read-only). Settings: `track_revisions`, `odd_and_even_pages_header_footer` (get and set) |
| `content_controls()` → tuple of `ContentControl` (`tag`, `alias`, `id`, `type`, `text`), `set_content_control_value_by_tag(tag, value)`, `set_content_control_value_by_alias(alias, value)` → count, `fill_content_control_tags([(tag, value), ...])` → counts | body controls; `type` such as `plain_text`, `rich_text`, `dropdown_list`, `combo_box`, `date`, `checkbox` (`"true"` / `"false"` ticks a box). Check the count: 0 means no control has that tag. The placeholder flag stays, a date gets no `w:fullDate` and a drop-down takes any text (gap docx-content-control-value-like-word) |
| `bind_content_controls({key: value, ...})` → count, `legacy_form_fields()` → tuple of `LegacyFormField` (`name`, `type` `"text"` / `"checkbox"` / `"dropdown"`, `value`, `choices`, `enabled`, `calculate_on_exit`, `source_part`, `ordinal`), `set_legacy_form_field_value(name, value, *, source_part=None, ordinal=None)` → `LegacyFormField` | `bind_content_controls` sets body controls by tag or alias in one call, as `rdocx fill --values` does (0 when no key matches: check the count). Legacy fields are Word's older form fields (`FORMTEXT`, `FORMCHECKBOX`, `FORMDROPDOWN`): a text field takes a str, a check box a bool, a drop-down the index of one of its `choices`; an unknown name raises `ValueError`, an index out of range `XmlError`; `source_part` and `ordinal` pick one of several fields sharing a name |
| `set_hyperlink_url(hyperlink, url)`, `remove_hyperlink(hyperlink)` | `hyperlink` from `doc.hyperlinks`. Removal keeps the text. After a removal or an added link, the older entries of that story raise `RdocxError`: re-fetch `doc.hyperlinks` |
| `image_data(rid)`, `replace_image(rid, bytes)`, `replace_image_for_story(story, rid, bytes)` | the relationship id is the `r:embed` of the picture (`Picture.relationship_id`). `replace_image` keeps the old extent |
| `add_comment(RunRange or StoryRunRange, *, author, text, initials=None, date=None)` → id | `RunRange(start=RunPosition(body_index=, run_index=), end=...)`, end exclusive, body paragraphs; `StoryRunRange(start=StoryRunPosition(item=, run_index=), end=...)` for a table cell or another story's paragraph item, or `StoryRunPosition(paragraph=p, run_index=)` from a `Paragraph` handle. No `date`, no date in the file: pass an RFC 3339 string. Ids stay stable across saves |
| `add_comment_on_text(anchor, *, author, text, occurrence=0, initials=None, date=None)` → id | anchors on exactly `anchor`, splitting runs with their format kept, table cells included: `occurrence` counts from 0 over the main story in document order, body and table cells together; an out-of-range occurrence raises `RdocxError`. `docx_ops.comment_on_text(..., in_tables=True)` counts cells only and checks the result |
| `reply_to(parent_id, *, author, text, date=None)` → id, `resolve_comment(id, *, resolved=True)`, `remove_comment(id)` | |
| `move_comment_to_text(id, anchor, *, occurrence=0)`, `move_comment(id, range)` | moves a thread onto exactly `anchor` (occurrences counted as `add_comment_on_text` does) or onto a `StoryRunRange` (a `Comment.anchor` read back works), keeping its id, author, date, text, replies and resolved flag; an emptied Google `goog_rdk` wrapper goes. A reply id, an unknown id or a text not found raises `RdocxError` |
| `accept_all()`, `reject_all()`, `accept_revision_id(id)`, `reject_revision_id(id)`, `accept_revisions_by_author(a)`, `reject_revisions_by_author(a)`, `accept_revisions_in_date_range(*, start, end)`, `reject_revisions_in_date_range(*, start, end)` | every supported story; dates as RFC 3339 strings |
| `compare(edited, author, timestamp, *, granularity="run", ignore_comments=False, ignore_formatting=False, ignore_whitespace=False, ignore_fields=False, ignored_stories=None)` → diagnostics (`ComparisonDiagnostic`: `location`, `message`) | turns `doc` into the redline of `doc` → `edited`; `granularity="word"` marks only the changed words (one deletion then one insertion per changed stretch), `ignore_comments=True` keeps `doc`'s comments and compares the rest, `ignored_stories=["header", "footer"]` keeps those stories of `doc`. A chart whose data changed shows as the old chart deleted and the new one inserted: accepting gives the edited chart, rejecting the original (both chart parts stay in the package) |
| `rebuild_toc()` → `TocRebuildReport` (`entry_count`, `bookmark_count`, `diagnostics`, `diagnostic_count`) | the entry of a numbered heading gets a left stop after its number |
| `insert_toc(bi, max_level=3)` | inserts a TOC field (`TOC \o "1-N" \h`) at body index `bi`; `rebuild_toc()` then fills it with entries linked to the headings, their page numbers and the `TOC1`... styles |
| `update_page_fields()` → int, `update_layout_backed_fields()` → `LayoutBackedFieldUpdateReport` (`page_fields`, `num_pages_fields`, `page_reference_fields`, `section_fields`, `section_pages_fields`, `updated_count`, `diagnostics`), `update_fields(*, now=datetime, file_name=, file_path=, merge_fields=, ...)` → count, `update_all_fields(...)` → `FieldsUpdateReport` (`ordinary_updated`, `layout`, `evaluations`: `FieldEvaluation` with `instruction`, `cached_result`, `outcome`; `total_updated`, `diagnostics`), `update_fields_on_open` (get/set) | caches field results from rdocx's pagination; PAGE and NUMPAGES in headers and footers keep their cached results, as Word does on save, and render with the right number anyway. `update_all_fields` is what `rdocx fields update` runs |
| `layout(*, font_dir=None)` → tuple of `LayoutFragment` (`body_index`, `physical_page`, `displayed_page`, `bounds`: a `BoundingBox`, `x`, `y`, `width`, `height` in points), `layout_page(i)` → `LayoutPage` (`page_number`, `displayed_page_number`, `width`, `height`) | rdocx's pagination, on Word's line heights |
| `to_pdf(*, font_dir=, fonts=, revision_view="accepted", revision_style="plain")` → bytes, `render_pages(*, dpi=150, format="png", quality=90, transparent=False, pages=None, font_dir=, revision_view=, revision_style=)` → list of bytes, `render_page_to_png(i, dpi, *, font_dir=, revision_view=, revision_style=)`, `render_all_pages(dpi, ...)`, `render_page_to_svg(i)` → `SvgRenderResult` (`svg`, `diagnostics`: `SvgDiagnostic` `path`, `message`) or None past the last page, `to_pdfa_deterministic(profile="pdfa-2b")` → bytes | `pages` zero-based. `revision_view="tracked"` shows tracked changes (deletions struck through, insertions underlined, a change bar); `revision_style="legal"` with it colours them (insertions blue with a double underline, deletions red, moves green; `ValueError` with the accepted view). `font_dir` puts a folder's fonts first, `font_dirs=[...]` several, `fonts=[(name, data)]` explicit faces. Each has a `..._with_diagnostics` twin returning `(result, diagnostics)` |

## Charts

`doc.add_chart(chart_type, chart_data, width, height, *, after=None)` → the `Paragraph` holding the chart (its
`text` is `""`) adds an editable chart with its embedded workbook, at the end of the body or after a body
`Paragraph`, from python-pptx's data objects, at the top level of `rdocx`:
`rdocx.CategoryChartData(number_format="General")` (`categories = [...]`, `add_series(name, values,
number_format=None)` → a `CategorySeriesData` with `name`, `values`, `number_format`) or `rdocx.XyChartData`
(`add_series(name)` → an `XySeriesData`; `add_data_point(x, y)`, `x_values`, `y_values`; every series with the
same x values). `chart_type` is a `rdocx.XL_CHART_TYPE`: `COLUMN_CLUSTERED`, `LINE`, `PIE`, `DOUGHNUT`, `AREA`,
`RADAR`, `XY_SCATTER`. Invalid data raises and changes nothing (one number format per chart).
`doc.chart_legend(paragraph)` → `{"visible": bool, "position": "r", "overlay": bool}` and
`doc.set_chart_legend(paragraph, *, visible=None, position=None, overlay=None)` read and set the legend of the
chart in that paragraph (positions `r`, `l`, `t`, `b`, `tr`; a new chart has none: `visible=True` adds it). A
Word document's charts are not read back (gap docx-chart-read-back) and their data is not replaced (rpptx's
`Chart.replace_data` is for decks).

## Fonts

| Member | Notes |
|---|---|
| `fonts` → tuple of `FontDefinition` (`name`, `alternate_name`, `family`, `pitch`, `embedded_fonts`: `EmbeddedFont` with `kind` `"regular"` / `"bold"` / `"italic"` / `"bold_italic"`, `data` (the font bytes, deobfuscated), `font_key`, `subsetted`, `license`) | the font table (`word/fontTable.xml`), a snapshot |
| `embed_font(name, font, kind="regular", *, license, font_key=None, subsetted=False)` | `font`: a path, bytes or a binary file of one TrueType or OpenType face (a `.ttc` collection raises `RdocxError`). `license=rdocx.FontEmbeddingLicense(authorized=True, identity="OFL-1.1: Example Sans")`: both keywords required; `authorized=False` or an empty identity raises `ValueError` and changes nothing. Creates the record when missing, obfuscates the face as Word does, and turns `embed_true_type_fonts` on. rdocx's own renders (`to_pdf`, PNG, `rdocx convert`) draw the embedded face; a `font_dir=` / `--font-dir` font of that name wins over it. Embed only a font whose licence allows it (SKILL.md, rule 9) |
| `remove_embedded_font(name, kind="regular")` → `EmbeddedFont` or None, `remove_font(name)` → `FontDefinition` or None, `set_font(rdocx.FontDefinition(name, alternate_name=None, family=None, pitch=None))` | remove one face (the record stays), a record and its faces, or add or replace a record (its faces kept). `family`: `roman`, `swiss`, `modern`, `script`, `decorative`, `auto`; `pitch`: `default`, `fixed`, `variable` |
| `embed_true_type_fonts`, `embed_system_fonts`, `save_subset_fonts` (get and set: True, False or None for unset) | the settings Word follows when it saves the document again: keep embedding the fonts, also embed common system fonts, embed only the characters in use (`w:embedTrueTypeFonts`, `w:embedSystemFonts`, `w:saveSubsetFonts`) |

Font records and faces are immutable detached snapshots. Invalid inputs and read failures leave the document
unchanged. Own rendering treats an absent embedding flag as enabled and explicit false as disabled.

Word can display only one embedded face per font name. Native Word may synthesize bold and discard the saved
bold part. The API retains all four face kinds and rdocx's renderer selects the supplied faces. Verify the
delivered file in Word when multiple faces of one family matter.

## Paragraph, Run, Font, ParagraphFormat

- `Paragraph`: `text` (the accepted view: tracked insertions in, deletions out. Setting it leaves one run
  without direct formatting, keeps the paragraph style, format and comment ranges, and refuses a
  paragraph holding one end of a field that spans paragraphs, such as a TOC), `runs`, `style` (reads the
  style id. Set an id or a name the document defines, as python-docx does: `KeyError` when no paragraph
  style has it, `ValueError` for a character or table style; `None` removes it), `alignment`
  (`WD_ALIGN_PARAGRAPH`), `numbering` (`(num_id, level)` or None, settable, unchecked), `paragraph_format`,
  `add_run(text)`, `replace_text(old, new, *, expect=None)` → int (counted replacement in this paragraph only,
  body or table cell: across runs, comments kept; a count other than `expect` raises `ReplacementCountError`
  and changes nothing), `xml` / `replace_xml(xml)` ("Raw XML").
- `add_hyperlink(text, url, *, tooltip=None)` → `Run` appends a web link (a `url` starting with `#` names a
  bookmark); `add_internal_hyperlink(text, bookmark, *, tooltip=None)` links to a bookmark name (an unknown
  one raises `RdocxError`); `doc.add_heading_hyperlink(paragraph, heading, text, *, tooltip=None)` appends to
  `paragraph` a link to the `heading` paragraph, which gets a `Heading_...` bookmark when it has none. The
  tooltip is written, not read back (`Hyperlink` has no `tooltip`).
- `add_equation(source, *, format)` → `EquationInsertResult` (`equation`, `diagnostics`) appends an inline
  equation from LaTeX (`format="latex"`) or MathML (`format="mathml"`); `RdocxError` when it does not parse.
  The LaTeX subset keeps fractions, radicals, scripts on their own base, `\pi`, `\infty`, `\to`, upright
  `\lim` and `\,` thin spaces; what it drops is a diagnostic. `equations()` → `Equation` (`is_display`,
  `to_latex()`, `to_mathml()` → an `EquationConversionResult` whose `equations` are `MathConversionResult`s:
  `equations[0].value` is the text, normalized (`E = mc^2` reads back `E=m{c}^{2}`), with
  `MathConversionDiagnostic`s (`path`, `message`) for what did not convert). Body and table-cell paragraphs. No display equation (gap docx-display-equation).
- `ParagraphFormat`: `space_before`, `space_after`, `line_spacing` (a float is a multiple, `2.0` double
  spacing; a length is exact), `left_indent`, `right_indent`, `first_line_indent`, `keep_with_next`,
  `keep_together`, `page_break_before`, `widow_control`, `alignment`, `outline_level` (0 to 9), `rtl`,
  `shading` (fill), `set_border(edge, style, *, size, color)`, `border(edge)` → `(style, size, RGBColor)` or
  None, `clear_border(edge)`, `clear_borders()` (edges `top`, `bottom`, `left`, `right`, `between`, `bar`),
  `tab_stops` (`add_tab_stop(position, alignment=WD_TAB_ALIGNMENT.LEFT, leader=WD_TAB_LEADER.SPACES)` keeps
  position order; `tab_stops[i]` is a read-only `TabStop` with `position`, `alignment`, `leader`; `del
  tab_stops[i]`, `clear_all()`, `len`). None means inherited; assigning None clears a direct value.
- `Run`: `text` (settable), `font`, `style_id`, `remove()`, `add_tab()` (a tab at the end of the run),
  `add_break(kind="line")` (`"page"`, `"column"`), `add_field(instruction, cached_result="")` (a simple field
  after the run, such as `"PAGE"`, showing `cached_result` until fields are updated), `add_picture(image,
  filename=None, width=None, height=None, *, options=None)` → `Picture` (the options of
  `Document.add_picture`; in a body, table-cell, header or footer run, a header or footer picture related from
  its own part), `xml` / `replace_xml(xml)`.
- `Font`: `name`, `size` (length), `bold`, `italic`, `underline` (`WD_UNDERLINE` or bool), `strike`,
  `double_strike`, `all_caps`, `small_caps` (written, not drawn by rdocx's renderer), `hidden`,
  `superscript` / `subscript` (python-docx's tri-state: setting one False clears only itself),
  `character_spacing` (a length, negative condenses), `highlight` (Word colour name: `"yellow"`, `"green"`,
  ..., `"none"`), `language`, `language_east_asia`, `language_bidi` (tags such as `"fr-FR"`, unchecked: gap docx-invalid-values-accepted),
  `name_east_asia`, `name_complex_script` (fonts that `name` also sets), `rtl`, `color` (a `ColorFormat`:
  `rgb`, `type` `"rgb"` / `"theme"` / `"auto"` or None, `theme_color` read as a string such as `"accent1"`,
  set as a string or `rdocx.MSO_THEME_COLOR.ACCENT_1`; `font.color = ...` is a shortcut for
  `font.color.rgb = ...`, None removes the colour), `shading` (fill, `RGBColor` or None). None means
  inherited, and assigning None clears the run's own value (bold, size, highlight...), so it takes the style's
  again.

## Tables

`Table`: `rows`, `cell(r, c)`, `add_row()` → `Row` (its `cells` can be written one after the other),
`insert_column(i, *, before=True)` (a column before, or after, column `i`), `delete_column(i)`,
`clone_row(index, at=None)` → `Row` (keeps cell and run formatting), `remove_row(index)` (a comment inside the
row goes with it), `style` (an id or a name the document defines, else `KeyError`), `alignment`
(`WD_TABLE_ALIGNMENT`), `width` (settable), `indent` (a length, negative pulls the table left, None removes
it), `autofit` (False writes a fixed layout), `first_row`, `last_row`, `first_column`, `last_column`,
`banded_rows`, `banded_columns` (the style's conditional formats), `xml` / `replace_xml(xml)`.
`Row.cells` (a merged cell appears once: use `cells[-1]` for the last column); `Cell`: `text` (settable),
`replace_text(old, new, *, expect=None)` → int (as `Paragraph.replace_text`, over the cell), `paragraphs`,
`add_paragraph(text)`, `add_table(rows, cols)` (a nested table sized to the cell) and `tables`,
`add_picture(...)`, `merge(other)` → the merged `Cell`, `split()` (undoes a horizontal merge), `width`
(settable: writes the cell width, not the table grid), `vertical_alignment` (`WD_CELL_VERTICAL_ALIGNMENT`),
`shading` (fill, settable, reads an `RGBColor`), `grid_span`, `vertical_merge`, `border(edge)` and
`set_border(edge, style, *, size, color)` (this cell), `margins` and `set_margins(*, top, right, bottom,
left)`, `xml` / `replace_xml(xml)`. On the table: `set_cell_grid_span(r, c, n)` (consumes empty cells only),
`set_cell_vertical_merge(r, c, "restart" | "continue" | None)`, `set_borders(style, *, size, color)`,
`set_border(edge, style, *, size, color)`, `border(edge)` → `(style, size, RGBColor)` or None,
`set_cell_margins(*, top, right, bottom, left)`, `cell_margins` → `(top, right, bottom, left)` or None,
`set_column_width(c, w)` and `grid_widths` (the grid, get and set). Edges: `top`, `bottom`, `left`, `right`,
`insideH`, `insideV`; styles: `none`, `single`, `thick`, `double`, `dotted`, `dashed`, `dotDash`, `wave`;
`size` in eighths of a point. `Row`: `height`, `height_rule` (`WD_ROW_HEIGHT_RULE.EXACTLY` or `AT_LEAST`),
`is_header` (repeated at the top of each page), `cant_split` (settable). A nested table has the same members, its paragraphs and XML included.

## Headers and footers

`doc.section(i).header`, `footer`, `first_page_header`, `first_page_footer`, `even_page_header`,
`even_page_footer` are live `HeaderFooter` handles; re-fetch them from `doc.section(i)` after an edit:

| Member | Notes |
|---|---|
| `paragraphs`, `add_paragraph(text="", style=None)` | a new header holds no paragraph: add one. A new paragraph takes the document's Header or Footer style by itself when it defines one; `style=` overrides it |
| `tables`, `add_table(rows, cols)` → `Table` | a header table is a `Table` (`cell(r, c)`, `rows`, `grid_widths` for its widths) |
| `is_linked_to_previous` (get and set) | setting False gives the section its own copy of the story it inherited, as Word does; writing into a linked one edits the earlier section's story |

Writing into a first-page story turns `different_first_page_header_footer` on for the section, into an even
story turns `doc.settings.odd_and_even_pages_header_footer` on (Word ignores even headers without it). Page
numbers are fields: `footer.add_paragraph("Page ").runs[0].add_field("PAGE", "1")`, then `add_run(" of ")`
and `add_field("NUMPAGES", "1")` on the paragraph re-fetched (`doc.section(0).footer.paragraphs[0]`); Word,
Google Docs and rdocx fill them on every page. A logo goes into a header or footer run with
`run.add_picture(...)`. The lower-level story calls (`create_section_story`, `insert_content`) stay
available.

## Raw XML

When a feature has no API, edit the element's XML instead of reaching for lxml: `Paragraph`, `Run`, `Table`
and `Cell` have `xml` (a str: the standalone element with its namespace declarations) and `replace_xml(xml)`
(a str, one element of the same kind), and `doc.section(i).xml` / `replace_xml(xml)` do the same for a
section's `w:sectPr`. Start from `xml` and change it: every prefix the replacement uses must be declared in it
(`xml` declares them on the root; a bare `<w:p>...` raises `XmlError`). A replacement is refused with an
`RdocxError` (`XmlError` for the XML itself), the document unchanged, when the XML is malformed or has a
DOCTYPE, holds another or a second element, leaves a cell without a paragraph, references a relationship id
the part lacks, or adds or removes a section break. It does not check the schema: an unknown `w:` element or
a foreign one is written (gap docx-raw-xml-unchecked). Body, table-cell, header and footer elements (a
header's relationship ids are those of its own part). A namespace declared on an ancestor inside the
replacement (on the `w:p` rather than on `asvg:svgBlip`) is kept. The element's own handle stays valid; the
handles inside it retire.

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
`column_spacing`, `different_first_page`, `break_type`, `page_number_start`; `doc.section(i)` for the live
handle); `Style` (`style_id`, `name`, `style_type`, `based_on`, `linked_style`, `next_style`, and the flags
`is_default`, `priority`, `hidden`, `semi_hidden`, `unhide_when_used`, `quick_format`, `locked`,
`auto_redefine`); `HeaderFooterVariant` (`section_index`, `kind` header/footer, `variant`
default/first/even, `story`, `inherited`, `source_section`); `Hyperlink` (`url`, `anchor`, `text`,
`index_path`, `relationship_id`, `story`); `Comment` (`id`, `author`, `initials`, `date`, `text`,
`parent_id`, `resolved`, `anchor_text`: the accepted-view text the comment covers, through Google `goog_rdk`
wrappers, paragraphs joined with `"\n"`, `""` for a reference without a range, None for a reply or a
comment with no markers; `anchor`: its `StoryRunRange` or None); `Revision` (`id`, `kind`, `author`,
`timestamp`, `story`); `StoryItem` (`kind` paragraph, table, drawing, content_control, field,
preserved_node; `index_path`, `direct_body_index`, `text`, `xml`, `story`, `revision`); `Story` (`kind`,
`part_name`, `owner_index`); `ContentLocation` (`story`, `item_kind`, `index_path`).
