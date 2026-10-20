# rpptx Python API (pinned build)

`import rpptx` with the pinned Python (`$R/python`). The API follows python-pptx in its main names; a python-pptx name
it lacks raises an error that names the rpptx way: `rpptx.Presentation`, `rpptx.util` (`Inches`, `Pt`, `Length` only:
there is no `Emu`, `Cm` or `Mm`; lengths are plain EMU integers, 914400 per inch, 12700 per point, and `Inches(1)` or
`Pt(12)` is a `Length`, an int with `.emu`, `.inches` and `.pt`), `rpptx.dml.color.RGBColor` (also
`RGBColor.from_string("7B1E3A")`; every colour setter also takes a hex string, with or without `#`, or an `(r, g, b)`
triple), `rpptx.enum.shapes` (`MSO_SHAPE`, `MSO_SHAPE_TYPE`, `MSO_CONNECTOR`, `PP_PLACEHOLDER`), `rpptx.enum.text`
(`PP_ALIGN`, `MSO_ANCHOR`, `MSO_AUTO_SIZE`, `MSO_UNDERLINE`), `rpptx.enum.dml` (`MSO_FILL`, `MSO_LINE_DASH_STYLE`,
`MSO_ARROWHEAD_STYLE`, `MSO_ARROWHEAD_WIDTH`, `MSO_ARROWHEAD_LENGTH`, `MSO_THEME_COLOR`, `MSO_COLOR_TYPE`,
`MSO_PATTERN_TYPE`), `rpptx.enum.lang` (`MSO_LANGUAGE_ID`), and at
the top level the chart names (`rpptx.XL_CHART_TYPE`, `rpptx.CategoryChartData`, `rpptx.XyChartData`). Errors:
`rpptx.RpptxError`, `XmlError`, `PackageError`, `StaleElementError`, `ReplacementCountError`. The modules have no
docstrings: exact signatures are in `rpptx/_rpptx.pyi` next to the installed module
(`$R/python -c "import rpptx, os; print(os.path.dirname(rpptx.__file__))"`).

## Handles

A handle (`Slide`, `Shape`, `TextFrame`, `Paragraph`, `Run`, a table's rows and cells, a layout) follows its own
element: moving, grouping or removing other shapes or slides, adding slides, importing slides, theme and font
changes leave it valid. It is retired, and raises `StaleElementError` naming the call and how to re-fetch it, when
its element is removed or rewritten:

| Edit | Retires |
|---|---|
| `shapes.remove(shape)`, `slides.remove(slide)`, `slide_layouts.remove(layout)` | that shape, slide or layout and what it holds |
| `text_frame.text = ...`, `shape.text = ...`, `text_frame.clear()`, a replacement that changes the frame (`try_replace_text` of the deck, slide or frame) | the paragraphs and runs of that frame (`paragraph.text = ...` the runs of that paragraph) |
| `element.replace_xml(xml)` | what lies below that element (a slide's: its shapes) |
| `rows.add_row`, `rows.remove`, `columns.add_column`, `columns.remove` | that table's rows and cells (the table's own shape handle stays) |

Everything else (setters, `run.text`, appends: `add_slide`, `add_textbox`, `add_chart`, `add_paragraph`,
`add_run`, `add_field`...) retires nothing: earlier paragraphs and runs stay live across appends, and only a
`paragraphs` or `runs` collection taken before an append raises (read `text_frame.paragraphs` or `p.runs` again).
`prs.slides` never goes stale. The handle a call returns (`group`, `ungroup`, `insert_picture`, `add_*`,
`duplicate`) is fresh.

## Presentation

| Member | Notes |
|---|---|
| `Presentation(path=None)` (a path), `Presentation.from_bytes(b)` (bytes, a stream's content for example), `save(path)` (a path), `to_bytes()` | `save` does not refuse the input (use `pptx_ops.save_atomic`); a new presentation is 16:9 with the eleven default layouts, and `save` writes the package class its extension names: a .potx saved as .pptx becomes a presentation; `save` to another format's extension (`"x.pdf"`) raises `ValueError`: write the bytes of `to_pdf()` |
| `slides`, `slide_layouts`, `slide_width`, `slide_height` (from 1 to 56 inches, `ValueError` otherwise), `slide_number_start` (the number of the first slide, None for 1), `comment_authors` (each a `CommentAuthor`: `id`, `name`, `initials`, `user_id`, `provider_id`) | |
| `slides.add_slide(layout)`, `slides.add_slide_at(layout, position)`, `slides.duplicate(slide)` → the copy, `slides.move(from_, to)`, `slides.remove(slide)`, `slide_layouts.index(layout)` | |
| `slides.import_slide(slide, layout=None, index=None, *, keep_source_formatting=False)` → the copy; `import_slide_with_diagnostics(...)` → (copy, diagnostics) | copies a slide of another presentation (at the end, or at `index`) with its charts, embedded objects, comments (authors merged) and notes; without `layout` it takes the destination layout of the same name and raises `RpptxError` ("no layout named ...") when there is none: pass `layout=prs.slide_layouts[k]`, or `keep_source_formatting=True` to bring the source layout, master and theme (each such import adds the source master, even an equal one; with `layout=` too it raises). Both presentations are read as they are in memory: a layout renamed or a theme colour changed before the import, in either, counts without saving first. A slide jump to a slide that is not copied is dropped and reported in the diagnostics: read them |
| `slides.import_slides(source, indexes=None, *, layout=None, index=None, keep_source_formatting=False, keep_sections=False)` → (copies, diagnostics) | several slides of `source` (zero-based `indexes`, all by default) in one call: a jump between two copied slides is kept. `keep_sections=True` brings their sections (the slides already here go into a "Default Section"). Merging two decks: `copies, diagnostics = prs.slides.import_slides(other)` |
| `try_replace_text(old, new, *, expect=None)` → int, `replace_all([(old, new, expect), (old, new)])` → tuple of counts | slides and notes, across runs, keeping the first run's formatting; a count other than `expect` raises `ReplacementCountError` and changes nothing (`replace_all`: all pairs or none). Without `expect`, a placeholder that is not found returns 0. `replace_text` is the same call as `try_replace_text` |
| `add_comment_author(*, id, name, user_id, provider_id, initials=None)` | id is a GUID in braces |
| `core_properties` (`CoreProperties`): `title`, `author`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `content_status`, `identifier`, `language`, `version` (str), `revision` (int), `created`, `modified`, `last_printed` (RFC 3339 strings such as `"2026-10-10T12:00:00Z"`) | read and write; `None` clears a field. A new presentation is stamped `2000-01-01T00:00:00Z`, revision 1: set `created` and `modified` when they matter |
| `app_properties` (`AppProperties`: `company`, `manager`, `template`, `application`, `slides`, `notes`, `hidden_slides`, `words`...), `custom_properties` (`CustomProperties`: `get(name)` → `CustomProperty` with `name`, `kind`, `value`; `set(name, value)` with a str, int, float, bool or `CustomFileTime("2026-10-10T12:00:00Z")`; `delete(name)`; iteration, `len`) | docProps/app.xml and docProps/custom.xml |
| `set_header_footer(slide_number=True, footer=None, date=None, hide_on_title=True, date_format="datetime1", *, date_cache=None)` | every slide, as PowerPoint's Header and Footer dialog with Apply to All: slide-owned `sldNum`, `ftr`, `dt` placeholders copied from the layout (or master), and the layouts' and masters' `p:hf` flags so later slides follow. `date`: None (off), `"auto"` (a `datetime1` to `datetime13` field that PowerPoint fills; empty until then unless `date_cache="16 October 2026"` gives its text) or fixed text; `hide_on_title` leaves title slides without them. Per slide: `slide.header_footer` |
| `slide_masters` (a sequence), `slide_master` (the first) | see Masters, layouts and themes |
| `apply_theme(source, *, source_master=None, import_masters=False, layout_map=None)` | colours and fonts of another deck or template (a path, or bytes: `other.to_bytes()`); `import_masters=True` also brings its masters and the layouts that match by type or name, keeping each slide's content and layout name |
| `sections` → tuple of `Section` (`id`, `name`, `slide_indices`, `slide_ids`), `set_sections([(name, [0, 1]), (name, [2])])` | `set_sections` takes every slide once, in slide order, else `ValueError` |
| `text_layout(*, width_factor=1.0, font_dir=None)` → tuple of `TextFrameLayout` (`slide_index`, `shape_id`, `name`, `overflow`, `autofit`, `font_scale`, `frame`, `usable`, `height`, `lines`: `TextLineLayout` with `text`, `font_size`, `baseline`, `bounds` (`BoundingBox`: `x`, `y`, `width`, `height`), `paragraph_index`; `diagnostics`) | rpptx's own line breaks, the slides' text frames (table cells: `fit_report`); `width_factor=0.95` asks whether text fits a narrower frame; percentage line spacing is laid out as LibreOffice does (100 % is 1.2 em) |
| `fit_report(*, font_dir=None)` → `FitReport` (`frames`: `FitFrameReport` with `context` (the part), `overflow`, `needed_scale` and `needed_reduction` (hundred-thousandths), `update` (an `AutofitUpdate`), `quantized_trial` and `persisted_trial` (`FitTrial`: `fits`, `font_scale`...); `diagnostics`) | what `rpptx fit` reports, every text owner of the slides, layouts, masters and notes, table cells included |
| `refresh_autofit(*, font_dir=None)` → `AutofitRefreshResult` (`updates`, `changed`, `diagnostics`) | writes the stored autofit of every frame whose `auto_size` asks for one: Text fit and autofit below |
| `validate()` → tuple of `ValidationIssue` (`kind`, `message`) | empty when the deck is valid |
| `sign(private_key, certificate)` → report, `verify_signatures()` → tuple of reports | a package signature over every part and relationship, written by the next `save`: the bytes of a PKCS#8 RSA key and an X.509 certificate (DER or one unencrypted PEM block each). A report is a dict: `signature_part`, `signer` (`subject`, `issuer`, `serial_number`), `cryptographically_valid`, `coverage_complete`, `covered_parts`, `covered_relationships`, `issues` (`kind` such as `changed_reference`, `details`); `verify_signatures()` checks the deck as it is now (an empty tuple when unsigned), so any edit after signing shows. A signed package or one with external relationships raises. Certificate trust is not checked (`references/cli.md`, Signatures) |
| `to_pdf(font_dir=None)`, `to_notes_pdf()`, `render_slide_to_png(i, dpi=150)`, `render_all_slides(dpi)`, `render_all_notes(dpi)`; `to_pdf_with_diagnostics()`, `render_slide_to_png_with_diagnostics(i)`... → (output, diagnostics) | zero-based slide index; the faces the deck embeds are drawn (not a compressed one, and the precedence of `font_dir=`: `embed_font`); `font_dir=` (also on the renders, `text_layout`, `fit_report` and `refresh_autofit`) adds fonts that are not installed. A slide with `show_master_shapes = False` is drawn without the master's shapes, a layout's own shapes stay |
| `render_slides(dpi=150, format="png", quality=None, transparent=False, slides=None)` | a list of PNG or JPEG images (`slides` zero-based), one multi-page TIFF as bytes for `format="tiff"`; `quality` with PNG or `transparent` with JPEG or TIFF raises `ValueError` |
| `to_pdfa("pdfa-2b")` (or `"pdfa-3b"`), `to_handout_pdf(slides_per_page=6)` (1, 2, 3, 4, 6 or 9), `render_all_handouts(slides_per_page, dpi)` | handouts with or without the deck's handout master; all accept keyword-only `fonts`, `font_dir` and `font_dirs` |
| `to_markdown(*, include_notes=False, skip_hidden=False, inline_images=True)`, `to_html(...)` → str; `export_markdown(...)`, `export_html(...)` → `PresentationTextExport` (`text`, `assets`: (file name, bytes) pairs, `diagnostics`); `save_markdown(path, ...)`, `save_html(path, ...)` → diagnostics | what `rpptx convert --to md` / `--to html` write: one section per slide headed by its title, in outline order, lists nested by level, tables, links, pictures with their alt text (data URIs when `inline_images`, else separate files that `save_*` writes next to `path` and refuses to overwrite); `include_notes=True` adds the speaker notes. The Markdown keeps whitespace exact: the spaces of titles, formatted runs and notes come out as `&#32;`, emphasis as `<strong>` / `<em>`; to read a deck's text, `rpptx text` or `outline` is plainer |
| `embedded_fonts` → tuple of `EmbeddedTypeface` (`typeface`, `panose`, `pitch_family`, `charset`, `faces`: `EmbeddedFontFace` with `kind`, `relationship_id`, `part_name`, `format` `"eot"`, `"compressed_eot"`, `"sfnt"` or `"unknown"`, `stored_len`, `data`: the font bytes, None when rpptx cannot decode the face) | the fonts the deck embeds (`p:embeddedFontLst`) |
| `embed_font(name, font, kind="regular", *, license)`, `remove_embedded_font(name, kind="regular")` → `EmbeddedFontFace` or None, `remove_font(name)` → `EmbeddedTypeface` or None | `font`: a path, bytes, binary reader or authored `EmbeddedFont` of one TrueType or OpenType face; `kind` `"regular"`, `"bold"`, `"italic"`, `"bold_italic"`; `license=rpptx.FontEmbeddingLicense(authorized=True, identity="OFL-1.1: Example Sans")` is required, and `authorized=False` or an empty identity makes `embed_font` raise `ValueError` and change nothing. The writer emits uncompressed EOT, retaining the font's own `fsType`, and turns on `embed_true_type_fonts`. Native PowerPoint checks retain uncompressed EOT and use genuine regular and bold faces. Verify the delivered deck in its receiving application. Own deterministic renders draw supported enabled embedded faces; a `font_dir=` font of that typeface wins over it, within the caller tier, exact sfnt families win before file aliases. A MicroType Express face (`format` `"compressed_eot"`, `data` None), which PowerPoint writes (observed upstream, tensorbee/rdocx#361), is not decoded: rpptx draws a substitute, so pass the font file with `font_dir=`. Embed only a font whose licence allows it (SKILL.md, rule 7) |
| `embed_true_type_fonts`, `save_subset_fonts` (get and set: True, False or None) | presence-sensitive `p:presentation` flags, None on a new deck until a font is embedded. Own rendering uses embedded faces only when `embed_true_type_fonts` is True. Verify font retention after resaving in the receiving application |
| `to_odp()` → (bytes, diagnostics), `save_odp(path)` → diagnostics, `Presentation.from_odp(path or bytes)` → (presentation, diagnostics) | OpenDocument; diagnostics list what was dropped as (path, reason): read them, a placeholder that inherits its geometry is dropped |

## Slide

`shapes`, `placeholders` (`slide.placeholders[idx]` looks one up by its `idx`; `len` and a loop give them all, in idx
order, a missing idx included), `slide_layout` (settable), `hidden` (settable),
`try_replace_text(old, new, *, expect=None, notes=True)` → int (this slide only, its notes unless `notes=False`;
same all-or-nothing count as the presentation's), `notes_text` (get and set; None when the slide has no notes,
setting it creates the notes slide), `has_notes_slide`, `notes_slide` (as python-pptx: creates the notes slide,
a `NotesSlide` whose `notes_text_frame` is a `NotesTextFrame` with `text`, get and set, or None),
`background.fill`, `follow_master_background`, `show_master_shapes` (False hides the master's and layout's shapes,
a logo for example), `comments` (each a `Comment` with `id`, `author_id`, `created`, `text`, `status`, `replies`:
`CommentReply` with the same fields but `replies`), `add_comment(*, id, author_id, created, text, shape_id=None,
text_start=None, text_length=None)` (on the slide, on the shape `shape_id` names, or on a character range of its
text; an unknown shape or author raises `RpptxError` and changes nothing),
`reply_to_comment(comment_id, *, id, author_id, created, text)`, `resolve_comment(comment_id)`,
`remove_comment(comment_id)`, `move_comment(from_, to)`, `move_reply(comment_id, from_, to)`. `created` is RFC 3339
with its zone. Two handles of one slide compare equal (`==`); a `Slide` is unhashable (no sets, no dict keys) and
`slide.slide_id` is the id PowerPoint registers (256 and up, as in python-pptx; `prs.sections[k].slide_ids` lists them
by section).

- `header_footer` (`HeaderFooter`: `slide_number` bool, `footer` text or None, `date` text, `"auto"` or None,
  `date_format`, settable; read only: `date_cache`, `date_mode`, `footer_present`, `date_present`,
  `slide_number_present`, `source` and `policy_source` (where the value comes from: `"slide"`, `"layout"`...)):
  read and write this slide's own placeholders, as `set_header_footer` does for all. A handle to a placeholder that
  a setter removes raises `StaleElementError`.
- `transition` (`SlideTransition`): `type` (`"fade"`, `"push"`, `"wipe"`, `"split"`, `"cover"`, `"uncover"`, `"cut"`,
  `"zoom"` or None), `direction` (checked per type: `"left"`, `"up"`, `"in"`...; another raises `RpptxError`),
  `duration` (seconds), `advance_on_click` (None until set), `advance_after` (seconds or None), `apply_to_all()`
  copies it to every slide.
- Media: `shapes.add_movie(file, left, top, width, height, poster_frame_image=None, mime_type=None)` (`mime_type`
  required when the file has no known extension), `slide.media` → tuple of `MediaInfo` (`shape_id`, `kind`
  `"video"` or `"audio"`, `content_type`, `part_name` for embedded media, `target` for a linked one, `volume`,
  `looped`, `trigger`, `trim_start_ms`...), `extract_media(shape_id)` → bytes, `replace_media(shape_id, file,
  mime_type=None)`, `remove_media(shape_id)` (the shape and its part).

## Masters, layouts and themes

- `prs.slide_master` / `prs.slide_masters[k]` (`SlideMaster`): `shapes` (the same `add_*` calls: a logo on every
  slide), `placeholders`, `background.fill`, `slide_layouts` (this master's, a tuple), `theme`, `text_styles`.
- `theme` (`Theme`): `name`, `colors` (`ThemeColors`, a mapping: `colors["accent1"]` → `RGBColor`, `colors["accent1"] =
  "7B1E3A"`, `keys()` (a tuple), iteration over the keys; keys `dk1`, `lt1`, `dk2`, `lt2`, `accent1` to `accent6`,
  `hlink`, `folHlink`, another raises `RpptxError`), `fonts` (`ThemeFonts`: `major` for headings, `minor` for body,
  each a `ThemeFontSet` with `latin`, `east_asian`, `complex_script`, settable).
- `text_styles` (`MasterTextStyles`): `title`, `body`, `other`, each a `TextStyleList` of nine levels
  (`body[0]` is level 1, like `paragraph.level`; also `body.level(1)`, one-based); a level has `font` (`name`, `size`, `bold`,
  `color`...), `bullet` (a character, or False), `bullet_color`, `left_indent`, `first_line_indent`.
- A layout (`prs.slide_layouts[k]`, `slide.slide_layout`, settable): `name` (settable), `slide_master`,
  `shapes`, `placeholders`, `background.fill`, `follow_master_background`, `show_master_shapes`,
  `used_by_slides` → tuple of `Slide`. By name: `next(lay for lay in prs.slide_layouts if lay.name == "Title Only")`
  (no `get_by_name`). `slide_layouts.duplicate(layout)` → the copy, under the same name (rename it),
  `slide_layouts.remove(layout)` (refused with `RpptxError` while a slide uses it).

## Shapes

- `ShapeCollection`: iteration, `title`, `placeholders`, `add_textbox(left, top, width, height)`,
  `add_shape(MSO_SHAPE.X, left, top, width, height)` (every preset of python-pptx's `MSO_SHAPE`, or a preset name
  such as `"roundRect"`), `add_connector(MSO_CONNECTOR.X, begin_x, begin_y, end_x, end_y)`, `add_picture(file, left,
  top, width=None, height=None)`, `add_table(rows, cols, left, top, width, height)`, `add_group_shape()` (an empty
  group), `remove(shape)`, `move(from_, to)` (z-order: index 0 is the back). A group's `shapes` take the same `add_*`
  calls, and the group grows to hold its members. `add_shape`, on a slide or in a group, writes python-pptx's
  `p:style` (accent1 fill and line, `effectRef idx="2"`, the minor font in `lt1`): the shape draws in the theme's
  colours with no direct fill or line (`fill.type` None) until they are set; `add_textbox` writes none.
  `shape.remove_style()` drops that `p:style` (an autoshape or a connector; a no-op without one).
  `add_connector` writes python-pptx's `p:style` (accent1 line, `effectRef idx="1"`: the theme's first effect style,
  an outer shadow in the default theme); `theme_effect_index = 0` gives a line without it.
- Arrange: `group([shapes])` → the group, members keep their slide position and id, a placeholder raises
  `RpptxError`; `shape.ungroup()` → the members, keeping their position; `align([shapes], "left"`, `"center"`,
  `"right"`, `"top"`, `"middle"` or `"bottom"`, `relative_to="selection"` or `"slide")`; `distribute([shapes],
  "horizontal"` or `"vertical", relative_to=...)` (at least three shapes relative to the selection).
- `Shape`: `shape_id`, `name` (settable), `shape_type` (`MSO_SHAPE_TYPE`), `left`, `top`, `width`, `height`
  (settable; None on a placeholder that inherits its geometry from the layout: `effective_geometry()` returns the
  inherited `(left, top, width, height)`, and a setter copies it before changing one value), `rotation`, `flip_h`,
  `flip_v` (`ValueError` on a table or chart frame), `has_text_frame`, `text` (setting it drops run
  formatting), `text_frame`, `fill`, `line`, `shadow`, `auto_shape_type` (get and set: `MSO_SHAPE` or a preset name
  such as `"roundRect"`; ValueError on a shape that is not an autoshape), `theme_effect_index` (the `effectRef` of
  `p:style`, None without one; setting it raises `RpptxError` on a shape without `p:style`), `click_action`
  (`hyperlink.address`: a web link on the whole shape, get and set; `target_slide`: a `Slide` or None, get and set, a
  click jumps to that slide; for a jump `hyperlink.address` reads the target part, such as `"slide3.xml"`; removing
  the target slide leaves a link that does nothing), `adjustments`, `has_table`, `table`, `image` (`blob`,
  `content_type`, `ext`), `replace_image(file)` (keeps position, size and crop), `crop_left`, `crop_top`,
  `crop_right`, `crop_bottom` (a picture's crop, fractions of its size, settable), `shapes` (a group's children),
  `xml` (bytes, read only) and `element` (Raw XML below).
- Charts: `shapes.add_chart(chart_type, left, top, width, height, chart_data)` → the graphic frame, an editable
  chart with its embedded workbook, from python-pptx's data objects: `rpptx.CategoryChartData` (`categories = [...]`,
  `add_series(name, values, number_format=None)` → `CategorySeriesData`: `name`, `values`, `number_format`) or
  `rpptx.XyChartData` (`add_series(name)` → `XySeriesData`: `add_data_point(x, y)`, `x_values`, `y_values`);
  `chart_type` a `rpptx.XL_CHART_TYPE` (`COLUMN_CLUSTERED`, `LINE` (with markers: python-pptx's `LINE_MARKERS`),
  `PIE`, `DOUGHNUT`, `AREA`, `RADAR`, `XY_SCATTER` (with lines)). `shape.has_chart`, `shape.chart` (`Chart`:
  `chart_type` (python-pptx's `XL_CHART_TYPE` value as an int, also for a stacked, horizontal, marker-free or
  exploded kind that rpptx does not write; None for a valid kind it cannot name);
  `categories`; `series`: `ChartSeries` with `name` and `values`, nan for a missing point; `has_legend`,
  `legend_position` (`"r"`, `"l"`, `"t"`, `"b"` or `"tr"`), `legend_overlay` (True draws the legend over the plot),
  get and set, drawn by the renders) reads any chart, `chart.replace_data(chart_data)` rewrites its numbers,
  categories and series names and its workbook, keeping each kept series' formatting.
- Accessibility: `alt_text` (`descr`, the text a screen reader and Google Slides read), `alt_title`, `decorative`
  (PowerPoint's "Mark as decorative"), on every kind of shape.
- Placeholders: `is_placeholder`, `placeholder_format` (`PlaceholderFormat`: `idx`, `type`: `PP_PLACEHOLDER`),
  `slide.placeholders[idx]`; `insert_picture(file)` on a picture or content placeholder → the picture, cropped to
  fill the frame.
- Connectors: `begin_connect(shape, site)`, `end_connect(shape, site)` glue an end to a connection site of
  another shape (for a rectangle 0 top, 1 left, 2 bottom, 3 right) and move it there; `begin_x`, `begin_y`,
  `end_x`, `end_y` read and move an end, releasing its glue.
- `TextFrame`: `paragraphs`, `add_paragraph()`, `text` (`""` leaves one paragraph with no run), `clear()` (as
  python-pptx: keeps the first paragraph and its formatting, drops its runs and the other paragraphs), `try_replace_text(old, new, *, expect=None)` → int
  (this frame only), `margin_left`, `margin_right`, `margin_top`, `margin_bottom`, `word_wrap`,
  `auto_size` (`MSO_AUTO_SIZE`), `autofit`, `vertical_anchor` (`MSO_ANCHOR`), `fit_text(...)`, `element`. Table
  cells have one too (`cell.text_frame`).
- `Paragraph`: `runs`, `add_run(text="")`, `text`, `alignment` (`PP_ALIGN`), `level`, `bullet` (a character, or
  False; also `bullet_char`), `bullet_color`, `bullet_size` (a multiple of the text size, or a
  `Length` such as `Pt(14)` for an absolute size), `bullet_font`,
  `auto_number` (a numbering scheme such as `"arabicPeriod"`, `"alphaLcParenR"`, `"romanUcPeriod"`),
  `auto_number_start`, `line_spacing` (a multiple such as 1.5 or a `Length`; zero or a negative one raises), `space_before`, `space_after` (a `Length` such as `Pt(6)`, or a float
  for a multiple of the line), `left_indent`, `right_indent`, `first_line_indent`, `font`,
  `add_field("slidenum")` (or `"datetime1"` to `"datetime13"`, optional cached text second; another type raises).
- `Run`: `text`, `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `all_caps`, `small_caps`,
  `baseline` (0.3 superscript, -0.25 subscript), `spacing` (character spacing, a `Length`: `Pt(2)`), `highlight_color`
  (a `ColorFormat`), `language` (a tag, `"fr-FR"`), `language_id` (the same as an `MSO_LANGUAGE_ID`, `FRENCH` for
  `"fr-FR"`; None clears it), `east_asian_name`, `complex_script_name`, `color`: a
  `ColorFormat` as in python-pptx, `font.color.rgb = ...` or the shortcut `font.color = ...`), `hyperlink.address`
  (get and set, also on a table cell's runs). Shape fills and lines use `fill.fore_color.rgb` and `line.color.rgb`, as in python-pptx.
- `ColorFormat`: `rgb`, `theme_color` (set `MSO_THEME_COLOR.ACCENT_2` or `"accent2"`: writes a scheme colour that
  follows the theme; reads an `MSO_THEME_COLOR`, `NOT_THEME_COLOR` for a direct RGB colour, and raises
  `AttributeError` when the run has no direct colour), `type` (`rpptx.enum.dml.MSO_COLOR_TYPE`: `RGB`, `SCHEME`,
  `HSL`, `PRESET`, `SCRGB`, `SYSTEM`, or None; compare with the enum, not a string; `MSO_THEME_COLOR_INDEX` is an alias of
  `MSO_THEME_COLOR`), `brightness` (-1 to 1).
- `FillFormat`: `solid()`, `background()` (no fill), `fore_color`, `type` (`MSO_FILL`), `transparency` (0 opaque to
  1, written as the colour's alpha, rounded as python-pptx does: 0.8 writes 20000); `gradient()` (two stops in accent1), `gradient_angle` (degrees,
  counter-clockwise as in python-pptx), `gradient_path` (None for linear), `gradient_stops` (`GradientStops`: index,
  `len`, `del gradient_stops[k]`, `append(position)` → a `GradientStop` with `position` 0 to 1 and `color`);
  `patterned()`, `pattern` (`MSO_PATTERN_TYPE`), `back_color`; `picture(file)` (shapes, table cells, backgrounds).
  `LineFormat`: `width`, `color.rgb`, `fill`, `dash_style` (`MSO_LINE_DASH_STYLE`), `head_end` and `tail_end`
  (`LineEndFormat`: `type` `MSO_ARROWHEAD_STYLE`, `width` `MSO_ARROWHEAD_WIDTH`, `length` `MSO_ARROWHEAD_LENGTH`).
- `ShadowFormat` (`shape.shadow`): `inherit` (False writes an empty `a:effectLst`: no theme shadow),
  `visible`, `color.rgb`, `alpha` (0 to 1), `blur_radius`, `distance` (EMU), `direction` (degrees), `align`
  (`"tl"` to `"br"`), `rotate_with_shape`; setting one writes `a:effectLst/a:outerShdw` in `spPr`.
- `Table`: `cell(r, c)` (a `Cell`: `text`, `text_frame`, `vertical_anchor`, `merge(other)`, `split()`,
  `is_merge_origin`, `is_spanned` (covered by a merge), `span_height`, `span_width`, `fill`, `margin_left`,
  `margin_right`, `margin_top`, `margin_bottom` (settable), `border_left`, `border_right`, `border_top`,
  `border_bottom` (each a `LineFormat`: set its `width` and `color.rgb`)), `columns[k].width`, `rows[k].height`
  (a render grows a row to its tallest cell text, as PowerPoint does; the file keeps the stored height),
  `rows[k].cells`, `rows.add_row(index=None)` → the new row, `rows.remove(row)`, `columns.add_column(index=None)`,
  `columns.remove(column)`: a new row or column copies a neighbour's size; re-fetch the cells and rows after each.
  Style options as in PowerPoint's Table Design tab: `first_row`, `last_row`, `first_col`, `last_col`,
  `horz_banding`, `vert_banding`; `style_id` (a built-in style GUID such as `"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"`,
  Medium Style 2 - Accent 1, or one of the deck's, stored as written; another raises `RpptxError`).

## Text fit and autofit

Setting `auto_size` writes the flag only: `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` ("shrink text on overflow") writes an
empty `a:normAutofit`, `SHAPE_TO_FIT_TEXT` an `a:spAutoFit`, and the shape keeps its size. `prs.refresh_autofit()`
then writes what PowerPoint stores: the `fontScale` and `lnSpcReduction` of each shrinking frame, the new height of
each resizing shape; call it after the last text change and before `save`, in either order of `auto_size` and text.
`text_frame.fit_text(font_family=None, max_size=18)` writes the largest whole point size up to `max_size` that fits
on every run (python-pptx's call; `font_family=None` keeps each run's font), turns autofit off and wrap on; it
raises `RpptxError` when the text does not fit even at 1 pt.

`AutofitUpdate` (from `fit_text`, in `refresh_autofit().updates` and `FitFrameReport.update`): `owner_kind`
(`"slide"`, `"layout"`, `"master"`, `"notes-master"`...), `slide_index`, `shape_id`, `name`, `cell`, `mode`
(`"normal"`, `"shape"`, `"none"`), `status` (`"updated"`, `"unchanged"`, `"empty"` for a frame without text),
`font_scale` and `line_spacing_reduction` (hundred-thousandths), `point_size` (from `fit_text`), `old_bounds` and
`new_bounds` (left, top, width, height in EMU), `changed`, `diagnostics` (the fonts measured with a stand-in, such as
`requested font "Calibri" measured with "Carlito"`: pass `font_dir=` or install them when the fit matters).

## Raw XML

The way out when rpptx has no API for an element (no lxml, no `_element`): `element` on `Shape`, `TextFrame`,
`Paragraph`, `Run`, `Table`, `Cell`, `Slide`, a layout and `SlideMaster` gives an `Element`; `element.xml` is the
element as a str (`p:sp`, `p:txBody`, `a:p`, `a:tbl`, `p:sld`...), `element.replace_xml(xml)` (a str) replaces it.
`shape.xml` is the same XML as bytes, read only. A replacement raises `RpptxError` and leaves the presentation
unchanged when the XML is malformed, has a DOCTYPE, has another root element or references an `r:id` the part lacks
(add the picture or link first), and so does an unknown or misplaced `p:` or `a:` element ("unknown or misplaced
direct XML child"): copy the structure from XML PowerPoint wrote. A replacement retires handles (Handles above).

## Errors instead of silent results

A slide whose groups (and `mc:AlternateContent`) nest deeper than 32 levels is refused at open with `XmlError`
("... nesting exceeds limit 32") instead of crashing, and `add_group_shape` refuses a 33rd level with `RpptxError`.

A bare int is EMU: `font.size = 12` (from 1 to 4000 points), `space_before = 6`, `space_after = 6` and
`font.spacing = 2` raise `ValueError` naming `Pt`, and so does `line_spacing = 0`: write `Pt(12)`, or a float for a
multiple of the line in spacing.
`font.name = ""` raises, and so do `save` to a `.pdf` path, a slide size outside 1 to 56 inches, a flip on a table or
chart frame and an unknown table `style_id`. An `RGBColor` is a value:
assign a new one (`fill.fore_color.rgb = "7B1E3A"`). `prs.slides._sldIdLst` raises an error naming
`prs.slides.remove(slide)` and `prs.slides.move(i, j)`; `shape._element` and `text_frame._txBody` raise
`NotImplementedError` naming `element.xml` and `element.replace_xml` (Raw XML).

Embedded typefaces and faces are immutable detached snapshots. Authored
`EmbeddedFont(font, kind="regular", *, license=...)` must match the kind and
licence passed to `embed_font`. Invalid inputs and read failures leave the deck
unchanged. Explicit `fonts=[]` uses deterministic embedded and bundled faces.
