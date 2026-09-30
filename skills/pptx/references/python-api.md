# rpptx Python API (pinned build)

`import rpptx` with the pinned Python (`$R/python`). The API follows python-pptx: `rpptx.Presentation`,
`rpptx.util` (`Inches`, `Pt`, `Length` only: there is no `Emu`, `Cm` or `Mm`; lengths are plain EMU
integers, 914400 per inch, 12700 per point), `rpptx.dml.color.RGBColor`, `rpptx.enum.shapes` (`MSO_SHAPE`,
`MSO_SHAPE_TYPE`, `MSO_CONNECTOR`), `rpptx.enum.text` (`PP_ALIGN`, `MSO_ANCHOR`, `MSO_AUTO_SIZE`,
`MSO_UNDERLINE`), `rpptx.enum.dml` (`MSO_FILL`, `MSO_LINE_DASH_STYLE`, `MSO_ARROWHEAD_STYLE`,
`MSO_ARROWHEAD_WIDTH`, `MSO_ARROWHEAD_LENGTH`). Errors: `rpptx.RpptxError`, `XmlError`, `PackageError`,
`StaleElementError`, `ReplacementCountError`. The modules have no docstrings: exact signatures are in `rpptx/_rpptx.pyi` next to the
installed module (`$R/python -c "import rpptx, os; print(os.path.dirname(rpptx.__file__))"`).

## Handles

`Slide`, `Shape`, `Paragraph`, `Run` handles are checked against the presentation's revision. Geometry,
font, paragraph and frame setters, fills, lines and line ends, shadows, `theme_effect_index`,
`auto_shape_type` and `slide.hidden` keep every handle valid. Any `add_*` call, `remove`, `move`, `add_slide`,
`import_slide` and every `try_replace_text` invalidate the handles of every slide. Setting `text_frame.text`,
`shape.text` or `notes_text` invalidates every handle, the shape an `add_*` just returned included;
`run.text` keeps them. Write `prs.slides[i].shapes[j]` again after each edit.

## Presentation

| Member | Notes |
|---|---|
| `Presentation(path=None)`, `Presentation.from_bytes(b)`, `save(path)`, `to_bytes()` | `save` does not refuse the input (use `pptx_ops.save_atomic`); a new presentation is 16:9 with the eleven default layouts, and `save` writes the package class its extension names: a .potx saved as .pptx becomes a presentation |
| `slides`, `slide_layouts`, `slide_width`, `slide_height`, `comment_authors` | |
| `slides.add_slide(layout)`, `slides.duplicate(slide)` → the copy, `slides.move(from_, to)`, `slides.remove(slide)`, `slide_layouts.index(layout)` | |
| `slides.import_slide(slide, layout=None, index=None)` → the copy | copies a slide of another presentation (at the end, or at `index`); without `layout` it takes the destination layout of the same name and raises `RpptxError` ("no layout named ...") when there is none: pass `layout=prs.slide_layouts[k]` |
| `try_replace_text(old, new, *, expect=None)` → int | slides and notes, across runs, keeping the first run's formatting; a count other than `expect` raises `ReplacementCountError` and changes nothing |
| `add_comment_author(*, id, name, user_id, provider_id, initials=None)` | id is a GUID in braces |
| `text_layout(*, width_factor=1.0)` → list of `TextFrameLayout` (`slide_index`, `shape_id`, `name`, `overflow`, `autofit`, `font_scale`, `frame`, `usable`, `height`, `lines`: `text`, `font_size`, `baseline`, `bounds`, `paragraph_index`) | rpptx's own line breaks; `width_factor=0.95` asks whether text fits a narrower frame; percentage line spacing is laid out as LibreOffice does (100 % is 1.2 em) |
| `to_pdf()`, `to_notes_pdf()`, `render_slide_to_png(i, dpi=150)`, `render_all_slides(dpi)`, `render_all_notes(dpi)` | zero-based slide index |

## Slide

`shapes`, `placeholders`, `slide_layout`, `hidden` (settable), `try_replace_text(old, new, *, expect=None,
notes=True)` → int (this slide only, its notes unless `notes=False`; same all-or-nothing count as the
presentation's), `notes_text` (None when the slide has no
notes; setting it creates the notes slide), `background.fill`, `follow_master_background`, `comments` (each with `id`, `author_id`, `created`,
`text`, `status`, `replies`), `add_comment(*, id, author_id, created, text)`, `reply_to_comment(comment_id,
*, id, author_id, created, text)`, `resolve_comment(comment_id)`, `remove_comment(comment_id)`, `move_comment(from_, to)`, `move_reply(comment_id, from_, to)`. `created`
is RFC 3339 with its zone.

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
- `Shape`: `shape_id`, `name`, `shape_type` (`MSO_SHAPE_TYPE`), `left`, `top`, `width`, `height` (settable;
  None on a placeholder that inherits its geometry from the layout: `effective_geometry()` returns the
  inherited `(left, top, width, height)`, and a setter copies it before changing one value), `rotation`, `has_text_frame`, `text` (setting it drops run
  formatting), `text_frame`, `fill`, `line`, `shadow`, `auto_shape_type` (get and set: `MSO_SHAPE` or a preset
  name such as `"roundRect"`; ValueError on a shape that is not an autoshape), `theme_effect_index` (the
  `effectRef` of `p:style`, None without one; setting it raises `RpptxError` on a shape without `p:style`),
  `adjustments`, `has_table`, `table`, `image` (`blob`,
  `content_type`, `ext`), `replace_image(file)` (keeps position, size and crop), `shapes` (a group's
  children), `xml` (the shape element as bytes).
- `TextFrame`: `paragraphs`, `add_paragraph()`, `text`, `try_replace_text(old, new, *, expect=None)` → int
  (this frame only), `margin_left/right/top/bottom`, `word_wrap`,
  `auto_size` (`MSO_AUTO_SIZE`), `autofit`, `vertical_anchor` (`MSO_ANCHOR`).
- `Paragraph`: `runs`, `add_run(text="")`, `text`, `alignment` (`PP_ALIGN`), `level`, `bullet`,
  `line_spacing`, `space_before`, `space_after`, `left_indent`, `right_indent`, `first_line_indent`, `font`.
- `Run`: `text`, `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `all_caps`, `color`: reads
  a hex string such as `"123456"`, set it with `font.color = RGBColor(...)`, not python-pptx's
  `font.color.rgb`), `hyperlink.address` (get and set; runs only). Shape fills and lines use
  `fill.fore_color.rgb` and `line.color.rgb`, as in python-pptx.
- `FillFormat`: `solid()`, `background()` (no fill), `fore_color.rgb`, `type` (`MSO_FILL`). `LineFormat`:
  `width`, `color.rgb`, `fill`, `dash_style` (`MSO_LINE_DASH_STYLE`), `head_end` and `tail_end`
  (`LineEndFormat`: `type` `MSO_ARROWHEAD_STYLE`, `width` `MSO_ARROWHEAD_WIDTH`, `length`
  `MSO_ARROWHEAD_LENGTH`).
- `ShadowFormat` (`shape.shadow`): `inherit` (False writes an empty `a:effectLst`: no theme shadow),
  `visible`, `color.rgb`, `alpha` (0 to 1), `blur_radius`, `distance` (EMU), `direction` (degrees), `align`
  (`"tl"` to `"br"`), `rotate_with_shape`; setting one writes `a:effectLst/a:outerShdw` in `spPr`.
- `Table`: `cell(r, c)` (`text`, `merge(other)`, `split()`, `is_merge_origin`, `span_height`, `span_width`,
  `fill`), `columns[k].width`, `rows[k].height`, `rows.add_row(index=None)` → the new row,
  `rows.remove(row)`, `columns.add_column(index=None)`, `columns.remove(column)`: a new row or column copies
  a neighbour's size. Re-fetch the table after each edit.
