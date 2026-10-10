# rpptx Python API (pinned build)

`import rpptx` with the pinned Python (`$R/python`). The API follows python-pptx: `rpptx.Presentation`,
`rpptx.util` (`Inches`, `Pt`, `Length` only: there is no `Emu`, `Cm` or `Mm`; lengths are plain EMU
integers, 914400 per inch, 12700 per point, and `Inches(1)` or `Pt(12)` is a `Length`, an int with `.emu`,
`.inches` and `.pt`), `rpptx.dml.color.RGBColor` (also `RGBColor.from_string("7B1E3A")`; every colour setter also takes a hex string, with or without `#`, or an `(r, g, b)` triple), `rpptx.enum.shapes` (`MSO_SHAPE`,
`MSO_SHAPE_TYPE`, `MSO_CONNECTOR`, `PP_PLACEHOLDER`), `rpptx.enum.text` (`PP_ALIGN`, `MSO_ANCHOR`, `MSO_AUTO_SIZE`,
`MSO_UNDERLINE`), `rpptx.enum.dml` (`MSO_FILL`, `MSO_LINE_DASH_STYLE`, `MSO_ARROWHEAD_STYLE`,
`MSO_ARROWHEAD_WIDTH`, `MSO_ARROWHEAD_LENGTH`, `MSO_COLOR_TYPE`, `MSO_THEME_COLOR`, `MSO_PATTERN_TYPE`),
`rpptx.enum.lang` (`MSO_LANGUAGE_ID`). Errors: `rpptx.RpptxError`, `XmlError`, `PackageError`,
`StaleElementError`, `ReplacementCountError`. The modules have no docstrings: exact signatures are in `rpptx/_rpptx.pyi` next to the
installed module (`$R/python -c "import rpptx, os; print(os.path.dirname(rpptx.__file__))"`).

## Handles

`Slide`, `Shape`, `Paragraph` and `Run` handles (table cells too) are checked: an edit retires the handles of
the kind it renumbers and those below it (slide > shape > paragraph > run), in every slide, and a retired
handle raises `StaleElementError` naming the call and how to re-fetch it. `prs.slides` never goes stale;
layout handles go stale only on the layout edits of the last row. The handle a call returns (`group`,
`ungroup`, `insert_picture`, `add_*`) is fresh.

| Edit | Retires |
|---|---|
| geometry, font, paragraph and frame setters, fills, lines, shadows, click actions, `run.text`, `notes_text`, `slide.hidden`, alt text, flips, `align`, `distribute`, `add_field`, `fit_text`, `refresh_autofit`, header and footer setters that add placeholders, transitions, theme, master text styles, core properties, sections; appends: `add_slide`, `add_textbox` and the other `add_*`, `add_paragraph`, `add_run` | nothing |
| `text_frame.text = ...`, `shape.text = ...` | paragraphs and runs |
| `try_replace_text` (deck, slide or frame) | runs |
| `shapes.remove`, `shapes.move`, `group`, `ungroup`, `insert_picture`, `Shape.replace_xml`, `TextFrame.replace_xml`, `gradient_stops.append` and `del gradient_stops[k]`, a header and footer call that removes a placeholder | shapes, paragraphs and runs |
| `slides.move`, `slides.remove`, `Slide.replace_xml` | slides and everything below |
| `slide_layouts.duplicate`, `slide_layouts.remove`, `layout.replace_xml`, `apply_theme` | layouts, shapes, paragraphs and runs |

## Presentation

| Member | Notes |
|---|---|
| `Presentation(path=None)` (a path or a binary stream), `Presentation.from_bytes(b)`, `save(path)` (a path or a stream), `to_bytes()` | `save` does not refuse the input (use `pptx_ops.save_atomic`); a new presentation is 16:9 with the eleven default layouts, and `save` writes the package class its extension names: a .potx saved as .pptx becomes a presentation; `save("x.pdf")` raises `ValueError` (write the bytes of `to_pdf()`) |
| `slides`, `slide_layouts`, `slide_width`, `slide_height` (1 to 56 inches, else `ValueError`), `comment_authors` (each a `CommentAuthor`: `id`, `name`, `initials`, `user_id`, `provider_id`) | |
| `slides.add_slide(layout)`, `slides.duplicate(slide)` → the copy, `slides.move(from_, to)`, `slides.remove(slide)`, `slide_layouts.index(layout)` | |
| `slides.import_slide(slide, layout=None, index=None)` → the copy | copies a slide of another presentation (at the end, or at `index`); without `layout` it takes the destination layout of the same name and raises `RpptxError` ("no layout named ...") when there is none: pass `layout=prs.slide_layouts[k]` |
| `try_replace_text(old, new, *, expect=None)` → int | slides and notes, across runs, keeping the first run's formatting; a count other than `expect` raises `ReplacementCountError` and changes nothing. `replace_text` is the same call |
| `add_comment_author(*, id, name, user_id, provider_id, initials=None)` | id is a GUID in braces |
| `core_properties` (`CoreProperties`): `title`, `author`, `subject`, `keywords`, `comments`, `category`, `last_modified_by`, `content_status`, `identifier`, `language`, `version` (str), `revision` (int), `created`, `modified`, `last_printed` (datetime, a naive one is read as UTC) | read and write; `None` clears a text field (it reads `""`). A new presentation carries none; nothing is stamped at save |
| `set_header_footer(slide_number=True, footer=None, date=None, hide_on_title=True, date_format="datetime1")` | every slide, as PowerPoint's Header and Footer dialog with Apply to All: slide-owned `sldNum`, `ftr`, `dt` placeholders copied from the layout (or master), and the layouts' and masters' `p:hf` flags so later slides follow. `date`: None (off), `"auto"` (a `datetime1` to `datetime13` field holding today's date, refreshed by PowerPoint) or fixed text; `hide_on_title` leaves title slides without them. Per slide: `slide.header_footer` |
| `slide_masters` (a sequence), `slide_master` (the first) | see Masters, layouts and themes |
| `apply_theme(source, *, import_master=False)` | colours and fonts of another deck or template (path, bytes, stream or `Presentation`); `import_master=True` also replaces the masters and the layouts that match by type or name, keeping each slide's content and layout name |
| `sections` → tuple of `Section` (`id`, `name`, `slide_indices`), `set_sections([(name, [0, 1]), (name, [2])])` | `set_sections` takes every slide once, in slide order, else `ValueError` |
| `text_layout(*, width_factor=1.0, font_dir=None)` → list of `TextFrameLayout` (`slide_index`, `shape_id`, `name`, `overflow`, `autofit`, `font_scale`, `frame`, `usable`, `height`, `lines`: `TextLineLayout` with `text`, `font_size`, `baseline`, `bounds` (`BoundingBox`: `x`, `y`, `width`, `height`), `paragraph_index`) | rpptx's own line breaks; `width_factor=0.95` asks whether text fits a narrower frame; percentage line spacing is laid out as LibreOffice does (100 % is 1.2 em) |
| `refresh_autofit()` → tuple of `AutofitResult` | writes the stored autofit of every frame that has one (Text fit and autofit below); `save`, `to_bytes`, the renders and `text_layout` already do it for the frames whose `auto_size` was set |
| `validate()` → tuple of `ValidationIssue` (`kind`, `message`) | empty when the deck is valid |
| `to_pdf(font_dir=None)`, `to_notes_pdf()`, `render_slide_to_png(i, dpi=150)`, `render_all_slides(dpi)`, `render_all_notes(dpi)` | zero-based slide index; `font_dir=` (also on the renders and `text_layout`) adds fonts that are not installed |
| `render_slides(dpi=150, format="png", quality=None, transparent=False, slides=None)` | a list of PNG or JPEG images (`slides` zero-based), one multi-page TIFF as bytes for `format="tiff"`; `quality` with PNG or `transparent` with JPEG or TIFF raises `ValueError` |
| `to_pdfa("pdfa-2b")` (or `"pdfa-3b"`), `to_handout_pdf(slides_per_page=6)` (1, 2, 3, 4, 6 or 9), `render_all_handouts(slides_per_page, dpi)` | handouts need the deck's handout master, else `RpptxError` saying how to get one (use `to_pdf()` instead); neither takes `font_dir` |
| `to_odp()` → (bytes, diagnostics), `save_odp(path)` → diagnostics, `Presentation.from_odp(path or bytes)` → (presentation, diagnostics) | OpenDocument; diagnostics list what was dropped as (path, reason): read them, a placeholder that inherits its geometry is dropped |

## Slide

`shapes`, `placeholders`, `slide_layout`, `slide_id`, `hidden` (settable), `try_replace_text(old, new, *, expect=None,
notes=True)` → int (this slide only, its notes unless `notes=False`; same all-or-nothing count as the
presentation's), `notes_text` (None when the slide has no
notes; setting it creates the notes slide), `has_notes_slide`, `notes_slide.notes_text_frame.text` (python-pptx's
spelling of `notes_text`, through a `NotesSlide` and its `NotesTextFrame`; reading `notes_slide` creates the notes slide), `background.fill`, `follow_master_background`, `show_master_shapes` (False hides the master's and
layout's shapes, a logo for example), `comments` (each a `Comment` with `id`, `author_id`, `created`,
`text`, `status`, `replies`: `CommentReply` with the same fields but `replies`), `add_comment(*, id, author_id, created, text, shape_id=None)` (on the slide, or on
the shape `shape_id` names; an unknown shape or author raises `RpptxError` and changes nothing),
`reply_to_comment(comment_id, *, id, author_id, created, text)`, `resolve_comment(comment_id)`, `remove_comment(comment_id)`, `move_comment(from_, to)`, `move_reply(comment_id, from_, to)`. `created`
is RFC 3339 with its zone. Two handles of one slide compare equal (`==`); a `Slide` is unhashable (no sets, no
dict keys).

- `header_footer` (`HeaderFooter`: `slide_number` bool, `footer` text or None, `date` text, `"auto"` or None,
  `date_format`): read and write this slide's own placeholders, as `set_header_footer` does for all. Re-fetch
  the shapes after switching one off: a handle to a removed placeholder raises `StaleElementError`.
- `transition` (`SlideTransition`): `type` (`"fade"`, `"push"`, `"wipe"`, `"split"`, `"cover"`, `"uncover"`, `"cut"`,
  `"zoom"` or None), `direction` (checked per type: `"left"`, `"up"`, `"in"`...), `duration` (seconds),
  `advance_on_click`, `advance_after` (seconds or None), `apply_to_all()` copies it to every slide.
- Media: `shapes.add_movie(file, left, top, width, height, poster_frame_image=None, mime_type=None)` (`mime_type`
  required when the file has no known extension), `slide.media` → tuple of `MediaInfo` (`shape_id`, `kind`
  `"video"` or `"audio"`, `content_type`, `linked`, `target`), `extract_media(shape_id)` → bytes,
  `replace_media(shape_id, file, mime_type=None)`, `remove_media(shape_id)` (the shape and its part).

## Masters, layouts and themes

- `prs.slide_master` / `prs.slide_masters[k]` (`SlideMaster`): `shapes` (the same `add_*` calls: a logo on every
  slide), `placeholders`, `background.fill`, `slide_layouts` (this master's), `theme`, `text_styles`.
- `theme` (`Theme`): `name`, `colors` (`ThemeColors`, a mapping: `colors["accent1"]` → `RGBColor`, `colors["accent1"] =
  "7B1E3A"`, `keys()`, `items()`; keys `dk1`, `lt1`, `dk2`, `lt2`, `accent1` to `accent6`, `hlink`, `folHlink`, another
  raises `KeyError`), `fonts` (`ThemeFonts`: `major` for headings, `minor` for body, each a `ThemeFontSet` with
  `latin`, `east_asian`, `complex_script`, settable).
- `text_styles` (`MasterTextStyles`): `title`, `body`, `other`, each a `TextStyleList` of nine levels
  (`body[0]` is level 1, like `paragraph.level`); a `TextStyleLevel` has `font` (`name`, `size`, `bold`,
  `color`...), `bullet` (a character, or False), `bullet_color`, `left_indent`, `first_line_indent`.
- A layout (`prs.slide_layouts[k]`, `slide.slide_layout`, settable): `name` (settable), `slide_master`,
  `shapes`, `placeholders`, `background.fill`, `follow_master_background`, `show_master_shapes`,
  `used_by_slides` → tuple of `Slide`. `slide_layouts.get_by_name(name, default=None)`,
  `slide_layouts.duplicate(layout)` → the copy (named `1_<name>`), `slide_layouts.remove(layout)` (refused
  with `ValueError` while a slide uses it).

## Shapes

- `ShapeCollection`: iteration, `title`, `placeholders`, `add_textbox(left, top, width, height)`,
  `add_shape(MSO_SHAPE.X, left, top, width, height)` (every preset of python-pptx's `MSO_SHAPE`),
  `add_connector(MSO_CONNECTOR.X, begin_x, begin_y, end_x, end_y)`, `add_picture(file, left, top,
  width=None, height=None)`, `add_table(rows, cols, left, top, width, height)`, `add_group_shape()`, `remove(shape)`, `move(from_, to)`
  (z-order: index 0 is the back). A group's `shapes` take the same `add_*` calls, and the group grows to
  hold its members. `add_shape`, on a slide or in a group, writes python-pptx's `p:style` (accent1 fill and
  line, `effectRef idx="2"`, the minor font in `lt1`): the shape draws in the theme's colours with no direct
  fill or line (`fill.type` None) until they are set; `add_textbox` writes none. `add_connector` writes python-pptx's `p:style` (accent1 line, `effectRef idx="1"`: the theme's first
  effect style, an outer shadow in the default theme); `theme_effect_index = 0` gives a line without it.
- Arrange: `group([shapes])` (also `add_group_shape([shapes])`) → the group, members keep their slide position and
  id, a placeholder raises `RpptxError`; `shape.ungroup()` → the members, keeping their position; `align([shapes], "left"`,
  `"center"`, `"right"`, `"top"`, `"middle"` or `"bottom"`, `relative_to="selection"` or `"slide")`;
  `distribute([shapes], "horizontal"` or `"vertical", relative_to=...)`.
- `Shape`: `shape_id`, `name`, `shape_type` (`MSO_SHAPE_TYPE`), `left`, `top`, `width`, `height` (settable;
  None on a placeholder that inherits its geometry from the layout: `effective_geometry()` returns the
  inherited `(left, top, width, height)`, and a setter copies it before changing one value), `rotation`, `flip_h`, `flip_v` (a table or chart raises `RpptxError`), `has_text_frame`, `text` (setting it drops run
  formatting), `text_frame`, `fill`, `line`, `shadow`, `auto_shape_type` (get and set: `MSO_SHAPE` or a preset
  name such as `"roundRect"`; ValueError on a shape that is not an autoshape), `theme_effect_index` (the
  `effectRef` of `p:style`, None without one; setting it raises `RpptxError` on a shape without `p:style`),
  `click_action` (`hyperlink.address`: a web link on the whole shape, get and set; `target_slide`: a `Slide` or
  None, get and set, a click jumps to that slide; for a jump `hyperlink.address` reads the target part, such as
  `"slide3.xml"`; removing the target slide leaves a link that does nothing), `adjustments`, `has_table`,
  `table`, `image` (`blob`, `content_type`, `ext`), `replace_image(file)` (keeps position, size and crop),
  `crop_left`, `crop_top`, `crop_right`, `crop_bottom` (a picture's crop, fractions of its size, settable), `shapes` (a group's
  children), `xml` and `replace_xml(xml)` (Raw XML below).
- Accessibility: `alt_text` (`descr`, the text a screen reader and Google Slides read), `alt_title`, `decorative`
  (PowerPoint's "Mark as decorative"), on every kind of shape.
- Placeholders: `is_placeholder`, `placeholder_format` (`PlaceholderFormat`: `idx`, `type`: `PP_PLACEHOLDER`), `slide.placeholders[idx]`;
  `insert_picture(file)` on a picture or content placeholder → the picture, cropped to fill the frame.
- Connectors: `begin_connect(shape, site)`, `end_connect(shape, site)` glue an end to a connection site of
  another shape (for a rectangle 0 top, 1 left, 2 bottom, 3 right) and move it there; `begin_x`, `begin_y`,
  `end_x`, `end_y` read and move an end, releasing its glue.
- `TextFrame`: `paragraphs`, `add_paragraph()`, `text`, `try_replace_text(old, new, *, expect=None)` → int
  (this frame only), `margin_left`, `margin_right`, `margin_top`, `margin_bottom`, `word_wrap`,
  `auto_size` (`MSO_AUTO_SIZE`), `autofit`, `vertical_anchor` (`MSO_ANCHOR`), `fit_text(...)`, `refresh_autofit()`,
  `xml` and `replace_xml(xml)`. Table cells have one too (`cell.text_frame`).
- `Paragraph`: `runs`, `add_run(text="")`, `text`, `alignment` (`PP_ALIGN`), `level`, `bullet` (a character, or
  False), `bullet_color`, `bullet_size` (a multiple of the text size), `bullet_font`, `auto_number` (a numbering
  scheme such as `"arabicPeriod"`, `"alphaLcParenR"`, `"romanUcPeriod"`), `auto_number_start`,
  `line_spacing` (a multiple such as 1.5 or a `Length`; zero or less raises), `space_before`, `space_after`, `left_indent`, `right_indent`, `first_line_indent`, `font`,
  `add_field("slidenum")` (or `"datetime1"` to `"datetime13"`, optional cached text second; another type raises).
- `Run`: `text`, `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `all_caps`, `small_caps`,
  `baseline` (0.3 superscript, -0.25 subscript), `spacing` (character spacing, a `Length`: `Pt(2)`; a bare int is EMU and one under 127 raises),
  `highlight_color`, `language` (a tag, `"fr-FR"`) or `language_id` (`MSO_LANGUAGE_ID.FRENCH`), `east_asian_name`,
  `complex_script_name`, `color`: a `ColorFormat` as in
  python-pptx, `font.color.rgb = ...` or the shortcut `font.color = ...`), `hyperlink.address` (get and set). Shape fills and lines use
  `fill.fore_color.rgb` and `line.color.rgb`, as in python-pptx.
- `ColorFormat`: `rgb`, `theme_color` (`MSO_THEME_COLOR.ACCENT_2`: writes a scheme colour that follows the theme),
  `type` (`MSO_COLOR_TYPE`), `alpha` (0 to 1, opacity).
- `FillFormat`: `solid()`, `background()` (no fill), `fore_color`, `type` (`MSO_FILL`); `gradient()` (two stops in
  accent1), `gradient_angle` (degrees, counter-clockwise as in python-pptx), `gradient_path` (None for linear),
  `gradient_stops` (`GradientStops`: index, `len`, `del gradient_stops[k]`, `append(position)` → a `GradientStop`
  with `position` 0 to 1 and `color`); `patterned()`, `pattern` (`MSO_PATTERN_TYPE`), `back_color`;
  `picture(file)` (shapes, table cells, backgrounds). `LineFormat`:
  `width`, `color.rgb`, `fill`, `dash_style` (`MSO_LINE_DASH_STYLE`), `head_end` and `tail_end`
  (`LineEndFormat`: `type` `MSO_ARROWHEAD_STYLE`, `width` `MSO_ARROWHEAD_WIDTH`, `length`
  `MSO_ARROWHEAD_LENGTH`).
- `ShadowFormat` (`shape.shadow`): `inherit` (False writes an empty `a:effectLst`: no theme shadow),
  `visible`, `color.rgb`, `alpha` (0 to 1), `blur_radius`, `distance` (EMU), `direction` (degrees), `align`
  (`"tl"` to `"br"`), `rotate_with_shape`; setting one writes `a:effectLst/a:outerShdw` in `spPr`.
- `Table`: `cell(r, c)` (a `Cell`: `text`, `text_frame`, `vertical_anchor`, `merge(other)`, `split()`, `is_merge_origin`, `is_spanned` (covered by a
  merge), `span_height`, `span_width`, `fill`, `margin_left`, `margin_right`, `margin_top`, `margin_bottom`
  (settable), `border_left`, `border_right`, `border_top`, `border_bottom` (each a `LineFormat`: set its
  `width` and `color.rgb`)), `columns[k].width`, `rows[k].height`, `rows[k].cells`, `rows.add_row(index=None)` → the new row,
  `rows.remove(row)`, `columns.add_column(index=None)`, `columns.remove(column)`: a new row or column copies
  a neighbour's size. Re-fetch the table after each edit. Style options as in PowerPoint's Table Design tab:
  `first_row`, `last_row`, `first_col`, `last_col`, `horz_banding`, `vert_banding`; `style_id` (a built-in style
  GUID such as `"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"`, Medium Style 2 - Accent 1, or one of the deck's; another raises `ValueError`).

## Text fit and autofit

`auto_size` is stored as PowerPoint stores it: `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` ("shrink text on overflow")
writes the `fontScale` and `lnSpcReduction` PowerPoint would compute, `SHAPE_TO_FIT_TEXT` resizes the shape
height to its text, at `save` / `to_bytes` / render / `text_layout` time, so setting `auto_size` before or
after the text gives the same file. `text_frame.fit_text(font_family=None, max_size=18, bold=False,
italic=False, font_file=None)` writes the largest whole point size up to `max_size` that fits on every run
(python-pptx's call; `font_family=None` keeps each run's font), turns autofit off and wrap on; it returns
None for an empty frame and raises `RpptxError` when the text does not fit even at 1 pt.

`AutofitResult` (from `fit_text`, `text_frame.refresh_autofit()` or `prs.refresh_autofit()`): `slide_index`,
`shape_id`, `name`, `autofit` (`"normal"`, `"shape"`, `"none"`), `font_scale`, `line_spacing_reduction`, `height`
and `width` (the new size when the shape was resized), `font_size` (from `fit_text`), `fits`,
`font_substitutions` (the (font, stand-in) pairs measured instead of fonts that are not installed: pass
`font_dir=` or install them when this is not empty and the fit matters).

## Raw XML

The way out when rpptx has no API for an element (no lxml, no `_element`: those raise `AttributeError` naming
this): `xml` gives bytes, `replace_xml(xml)` (str or bytes) replaces the element, on `Shape` (`p:sp`, `p:pic`,
`p:cxnSp`...: the same kind), `TextFrame` (`p:txBody`), `Slide` (`p:sld`) and a layout (`p:sldLayout`).
A replacement raises `ValueError` and leaves the presentation unchanged when the XML is malformed, has a
DOCTYPE, has another root element, references an `r:id` the part lacks (add the picture or link first), or puts
an unknown or misplaced `p:` or `a:` element anywhere in the slide, shape or text XML (`a:graphicData`, `p:timing`
and `p:transition` stay unchecked; `mc:AlternateContent` and `mc:Ignorable` content are allowed). Then run
`rpptx validate` and render. A replacement retires handles (Handles above).

## Errors instead of silent results

A bare int is EMU: `font.size = 12` raises `ValueError` (from 1 to 4000 points), and so does `space_before = 6`
or `space_after = 6` (under the 0.01 pt PowerPoint stores): write `Pt(12)`, or a float for a multiple of the
line in spacing. `font.name = ""` raises; `save("x.pdf")` raises; `RGBColor` is immutable (`.rgb =` raises
`AttributeError`: assign a new colour). A python-pptx spelling rpptx lacks (`shape._element`,
`prs.slides._sldIdLst`, `text_frame._txBody`) raises `AttributeError` naming the rpptx call.
