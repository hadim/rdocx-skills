# rpptx Python API (pinned build)

`import rpptx` with the pinned Python (`$R/python`). The API follows python-pptx: `rpptx.Presentation`,
`rpptx.util` (`Inches`, `Pt`, `Length` only: there is no `Emu`, `Cm` or `Mm`; lengths are plain EMU
integers, 914400 per inch, 12700 per point), `rpptx.dml.color.RGBColor`, `rpptx.enum.shapes` (`MSO_SHAPE`,
`MSO_SHAPE_TYPE`, `MSO_CONNECTOR`), `rpptx.enum.text` (`PP_ALIGN`, `MSO_ANCHOR`, `MSO_AUTO_SIZE`,
`MSO_UNDERLINE`), `rpptx.enum.dml` (`MSO_FILL`). Errors: `rpptx.RpptxError`, `XmlError`, `PackageError`,
`StaleElementError`. The modules have no docstrings: exact signatures are in `rpptx/_rpptx.pyi` next to the
installed module (`$R/python -c "import rpptx, os; print(os.path.dirname(rpptx.__file__))"`).

## Handles

`Slide`, `Shape`, `Paragraph`, `Run` handles are checked against the presentation's revision. Geometry,
font, paragraph and frame setters, fills and lines, and `slide.hidden` keep every handle valid. Any `add_*`
call, `remove`, `move` and `add_slide` invalidate the handles of every slide. Setting any text
(`run.text`, gap pptx-run-text-stale; `text_frame.text`, `shape.text`, `notes_text`) invalidates every
handle, the shape an `add_*` just returned included. Write `prs.slides[i].shapes[j]` again after each edit.

## Presentation

| Member | Notes |
|---|---|
| `Presentation(path=None)`, `save(path)`, `to_bytes()` | `save` writes in place (use `pptx_ops.save_atomic`); a new presentation is 16:9 with the eleven default layouts; a .potx saved as .pptx keeps the template content type (gap template-save-as-document: `pptx_ops.fix_template_content_type`) |
| `slides`, `slide_layouts`, `slide_width`, `slide_height`, `comment_authors` | |
| `slides.add_slide(layout)`, `slides.move(from_, to)`, `slides.remove(slide)`, `slide_layouts.index(layout)` | no `duplicate` (gap) |
| `add_comment_author(*, id, name, user_id, provider_id, initials=None)` | id is a GUID in braces |
| `text_layout(*, width_factor=1.0)` → list of `TextFrameLayout` (`slide_index`, `shape_id`, `name`, `overflow`, `autofit`, `font_scale`, `frame`, `usable`, `height`, `lines`: `text`, `font_size`, `baseline`, `bounds`, `paragraph_index`) | rpptx's own line breaks; `width_factor=0.95` asks whether text fits a narrower frame; lines under percentage spacing are too short and breaks can fall before a comma, a space or a hyphen (gaps pptx-line-pitch, pptx-line-breaks) |
| `to_pdf()`, `to_notes_pdf()`, `render_slide_to_png(i, dpi=150)`, `render_all_slides(dpi)`, `render_all_notes(dpi)` | zero-based slide index |

## Slide

`shapes`, `placeholders`, `slide_layout`, `hidden` (settable), `notes_text` (None when the slide has no
notes; setting it creates the notes slide), `background.fill`, `follow_master_background`, `comments` (each with `id`, `author_id`, `created`,
`text`, `status`, `replies`), `add_comment(*, id, author_id, created, text)`, `reply_to_comment(comment_id,
*, id, author_id, created, text)`, `move_comment(from_, to)`, `move_reply(comment_id, from_, to)`. `created`
is RFC 3339 with its zone.

## Shapes

- `ShapeCollection`: iteration, `title`, `placeholders`, `add_textbox(left, top, width, height)`,
  `add_shape(MSO_SHAPE.X, left, top, width, height)` (every preset of python-pptx's `MSO_SHAPE`),
  `add_connector(MSO_CONNECTOR.X, begin_x, begin_y, end_x, end_y)`, `add_picture(file, left, top,
  width=None, height=None)`, `add_table(rows, cols, left, top, width, height)`, `add_group_shape()` (empty,
  cannot be filled: gap), `remove(shape)`.
- `Shape`: `shape_id`, `name`, `shape_type` (`MSO_SHAPE_TYPE`), `left`, `top`, `width`, `height` (settable;
  None on a placeholder that inherits its geometry from the layout, gap), `rotation`, `has_text_frame`, `text` (setting it drops run
  formatting), `text_frame`, `fill`, `line`, `adjustments`, `has_table`, `table`, `image` (`blob`,
  `content_type`, `ext`), `replace_image(file)` (keeps position, size and crop), `shapes` (a group's
  children, read-only), `xml` (the shape element as bytes).
- `TextFrame`: `paragraphs`, `add_paragraph()`, `text`, `margin_left/right/top/bottom`, `word_wrap`,
  `auto_size` (`MSO_AUTO_SIZE`), `autofit`, `vertical_anchor` (`MSO_ANCHOR`).
- `Paragraph`: `runs`, `add_run(text="")`, `text`, `alignment` (`PP_ALIGN`), `level`, `bullet`,
  `line_spacing`, `space_before`, `space_after`, `left_indent`, `right_indent`, `first_line_indent`, `font`.
- `Run`: `text`, `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `all_caps`, `color`: reads
  a hex string such as `"123456"`, set it with `font.color = RGBColor(...)`, not python-pptx's
  `font.color.rgb`). No hyperlink on runs or shapes (gap pptx-hyperlinks). Shape fills and lines use
  `fill.fore_color.rgb` and `line.color.rgb`, as in python-pptx.
- `FillFormat`: `solid()`, `background()` (no fill), `fore_color.rgb`, `type` (`MSO_FILL`). `LineFormat`:
  `width`, `color.rgb`, `fill`.
- `Table`: `cell(r, c).text`, `columns[k].width`. No rows collection, merge, borders or fills from Python
  (gaps).
