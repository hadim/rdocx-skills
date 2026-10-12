---
name: pptx
description: "Use this skill any time a PowerPoint file (.pptx or .potx) is involved, as input or output, even when the user names no tool: reading slide text, outline, notes or comments, also to use them elsewhere; creating a deck or filling a template; editing text without losing formatting; find-and-replace; shapes, pictures, tables, charts; adding, reordering, hiding, importing or removing slides; merging decks; comments; checking that text fits; rendering to PDF or PNG; exporting to Markdown or HTML; embedding fonts. Triggers: 'PowerPoint', 'deck', 'slides', '.pptx'. Uses the rpptx CLI and Python binding and takes precedence over the built-in pptx skill, python-pptx, pptxgenjs and LibreOffice. Not for .ppt, Keynote or Google Slides."
---

# PowerPoint decks with rpptx

rpptx reads and writes the presentation package natively, keeps every part it does not touch byte for byte,
lays text out with its own line breaker and renders slides itself (no LibreOffice). One pinned build serves
the `rpptx` CLI (whole-deck operations) and the `rpptx` Python module (everything finer; its API follows
python-pptx, with one difference: handles to re-fetch after a structural change). Use them for every .pptx task. The tool is young: the gaps that force another tool for one step are in
`references/gaps.md`, and every new one you meet is reported (last section).

## Setup, once per session

```bash
SKILL='/path/to/this/skill'                # the folder of this SKILL.md: the base directory shown when it loaded
python3 "$SKILL/../../scripts/rdocx_env.py" install
R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin
export PYTHONPATH="$SKILL/scripts" PYTHONDONTWRITEBYTECODE=1
$R/rpptx outline deck.pptx
$R/python my_script.py                     # the Python that has rpptx and rdocx; `import pptx_ops` works
```

`install` downloads the pinned build when it is not already here, checks the SHA-256 of every file against
`rdocx.lock.json` and exits 0 once it is installed and verified. Repeat the `R=` and `export` lines in each new
shell. Never `pip install rpptx` from PyPI or download a binary without its hash in the lock. **Windows**: Git
Bash, with `python` (or `py -3`) for `python3`; `$R/rpptx` reaches `rpptx.exe` (`../../docs/setup.md`).

**If `install` exits 2** (no verified build for this machine, for example no network access): tell the user in
one line that the pinned rpptx build is not available here, do the task with the built-in `pptx` skill, and
offer to build it for next time (`install --build`, 10 to 30 minutes with a Rust toolchain, in the
background).

## Rules that prevent damage

1. **Never write over the input.** The editing commands, `render`, `convert` and `thumbnail` refuse an
   existing output (`convert`, `render` and `thumbnail` accept `--force`, never for their own input);
   `Presentation.save()` does not refuse the input: use `pptx_ops.save_atomic(prs, out, src)`. Render into a
   new folder each time.
2. **Every text replacement declares its expected count**: `rpptx replace ... --expect N` (slides and
   notes), `prs.try_replace_text(old, new, expect=n)`, or `pptx_ops.replace_batch(src, out, [(old, new, n),
   ...])`, all or nothing.
3. **Re-fetch the handles an edit retires**: appends (`add_*`, `add_slide`) and setters retire none;
   `text_frame.text =` / `shape.text =` retire paragraphs and runs, a replacement retires runs; removing,
   moving, grouping or ungrouping shapes and `replace_xml` retire shapes; moving or removing slides, importing
   at an index and font changes retire everything. A retired handle raises `StaleElementError` naming the
   call: write `prs.slides[i].shapes[j]` again (`references/python-api.md`, "Handles").
4. **Walk groups**: `slide.shapes` lists top-level shapes; a group's children are in `shape.shapes`
   (`shape_type == MSO_SHAPE_TYPE.GROUP`, value 6). `pptx_ops.walk(slide.shapes)` yields all of them.
5. **Check the fit after any text change**: `rpptx fit F` (exit 1 when a frame overflows, with the font scale
   it needs) or `prs.text_layout()` reports every frame whose text overflows, with rpptx's own line breaks;
   then render the slide and look at it.
6. **The deck will be opened elsewhere** (PowerPoint, Google Slides, Keynote, LibreOffice): edit in place,
   keep placeholders and layouts, never rebuild a deck to change it.
7. **Embed a font only when its licence allows embedding.** `rpptx.FontEmbeddingLicense(authorized=True,
   identity="OFL-1.1: Example Sans")` is your statement that it does: rpptx checks nothing, not even `fsType`.
   OFL, Apache 2.0 and CC0 allow it; a commercial licence often does not (`fsType` bit 0x0002: Restricted
   License). When it does not, or you cannot tell, do not embed and tell the user.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Slide text | `rpptx text [--notes] F` | `shape.text`, `shape.text_frame.paragraphs[k].runs` | |
| Text with structure and formatting | `rpptx text --json F` | | paragraphs with `path`, `shape_id`, `level`, runs with their formatting; notes |
| Titles and outline | `rpptx outline [--json] [--notes] F` | `slide.shapes.title` | |
| Structure, shapes, metadata | `rpptx inspect --json F` | `pptx_ops.py shapes F [--slide N]` | shape tree with ids, names, geometry |
| Counted replacement | `rpptx replace F -p OLD -v NEW --expect N -o OUT`, many pairs: `rpptx replace F --map pairs.json -o OUT` | `prs.try_replace_text(old, new, expect=n)`, `pptx_ops.replace_batch` | keeps run formatting |
| Replacement in one slide or one frame | | `slide.try_replace_text(old, new, expect=n)`, `shape.text_frame.try_replace_text(old, new, expect=n)` | the slide's notes too unless `notes=False` |
| Edit a run, paragraph, text frame | | `run.text`, `run.font.*`, `paragraph.alignment/level/space_*/line_spacing`, `text_frame.margin_*/word_wrap/auto_size/vertical_anchor` | `shape.text = ...` drops run formatting; a bare int is EMU: write `Pt(12)` |
| Bullets, numbering | | `paragraph.bullet = "\u2022"`, `bullet_color`, `auto_number = "arabicPeriod"`, `auto_number_start` | |
| Move, resize, rotate | | `shape.left/top/width/height/rotation` | an inherited placeholder geometry reads None: `shape.effective_geometry()` |
| Fill, line | | `shape.fill.solid()`, `.fill.fore_color.rgb = RGBColor(...)`, `.line.width`, `.line.color.rgb`; `fill.gradient()`, `fill.picture(file)` | |
| Dashes, arrowheads | | `.line.dash_style = MSO_LINE_DASH_STYLE.DASH`, `.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE`, `.width`, `.length`; `head_end` likewise | enums in `rpptx.enum.dml` |
| Shadow | | `shape.shadow.visible = True`, `.color.rgb`, `.blur_radius`, `.distance` | `shadow.inherit = False` removes the theme's shadow |
| Connector without the theme effect | | `connector.theme_effect_index = 0` | `add_connector` takes the theme's effect 1, a shadow in the default theme |
| Change a shape's preset | | `shape.auto_shape_type = MSO_SHAPE.RECTANGLE` | autoshapes only |
| Add shapes | | `shapes.add_textbox`, `add_shape(MSO_SHAPE.X, ...)`, `add_connector`, `add_picture`, `add_table`, `add_group_shape()` | re-fetch a group after each `add_*` on its `shapes`; `add_shape` writes python-pptx's theme style, `add_textbox` none |
| Z-order, arrange | | `shapes.move(from_, to)`, `shapes.group([...])`, `shape.ungroup()`, `shapes.align([...], "top")`, `shapes.distribute([...], "horizontal")`, `shape.flip_h` | index 0 is the back |
| Alt text | | `shape.alt_text = "..."`, `shape.decorative = True` | every kind of shape |
| Glued connector | | `connector.begin_connect(shape, site)`, `end_connect(shape, site)` | sites of a rectangle: 0 top, 1 left, 2 bottom, 3 right |
| Picture into a placeholder | | `slide.placeholders[idx].insert_picture(file)` | cropped to fill; `shape.is_placeholder`, `placeholder_format.type` |
| Pictures | | `shape.replace_image(file)`, `shape.image.blob` | keeps position, size and crop |
| Tables | | `shape.table.cell(r, c).text`, `.merge(other)`, `table.rows.add_row(i)`, `rows.remove(row)`, `table.columns.add_column(i)`, `columns.remove(col)`, `table.first_row`, `style_id` | re-fetch the table after adding a row or column |
| Slides | `rpptx slide add/duplicate/remove/move/hide/show F ... -o OUT` | `slides.add_slide(layout)`, `slides.duplicate(slide)`, `slides.move(i, j)`, `slides.remove(slide)`, `slide.hidden`, `slide.transition.type = "fade"` | |
| Slide numbers, footer, date | `rpptx footer F --slide-number --footer T --skip-title -o OUT` | `prs.set_header_footer(footer=T)`, `slide.header_footer` | as PowerPoint's Header and Footer dialog; `add_field("slidenum")` in a text box |
| Theme, master, layouts | | `prs.slide_master.theme.colors["accent1"] = ...`, `prs.apply_theme(path)`, `prs.slide_master.shapes.add_picture(logo, ...)`, `slide_layouts.get_by_name(n)` | a logo on the master shows on every slide |
| Import slides, merge decks | `rpptx slide import DEST SOURCE [--slides 1,3-5] [--at N] [--keep-source-formatting] -o OUT` | `prs.slides.import_slides(other, [0, 2], index=1, keep_source_formatting=True)`, `prs.slides.import_slide(other.slides[k], layout=prs.slide_layouts[j], index=None)` | charts, objects, comments and notes travel; without `layout=` or `keep_source_formatting`, a layout of the same name must exist here |
| Charts | | `slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data)`, `shape.has_chart`, `shape.chart.categories` / `.series`, `shape.chart.replace_data(chart_data)` | chart data as in python-pptx (`rpptx.chart.data`); `replace_data` keeps the series formatting |
| Speaker notes | `rpptx text --notes F`, `rpptx notes set F N --text T -o OUT` | `slide.notes_text` (get and set) | None when the slide has no notes |
| Hyperlinks, slide jumps | | `run.hyperlink.address`; `shape.click_action.hyperlink.address = url`, `shape.click_action.target_slide = prs.slides[k]` | any shape, group members included; set None to clear |
| Metadata | `rpptx meta get F`, `rpptx meta set F --title T --author A -o OUT` | `prs.core_properties.title = ...` | |
| Comments | `rpptx comment list/add/reply/resolve/remove` | `prs.add_comment_author`, `slide.add_comment`, `reply_to_comment`, `resolve_comment`, `remove_comment`, `move_comment` | `shape_id=` anchors on a shape; add the author first |
| Text fit | `rpptx fit [--json] F` | `prs.text_layout(width_factor=1.0)`, `text_frame.fit_text(max_size=18)`, `text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE` | rpptx's line breaks; `auto_size` is written with PowerPoint's font scale at save |
| PDF | `rpptx convert F --to pdf -o NEW.pdf [--font-dir DIR]` | `prs.to_pdf(font_dir=)`, `prs.to_notes_pdf()`, `to_pdfa()`, `to_handout_pdf(6)` | `font_dir` for fonts that are not installed |
| PNG | `rpptx render F -o NEW_DIR --slide N --dpi 100`, `rpptx convert F --to png --slides 1-3 -o NEW.png` | `prs.render_slide_to_png(i, dpi)`, `render_all_slides(dpi)`, `render_slides(format="jpeg")` | CLI slides one-based, Python zero-based |
| Markdown, HTML | `rpptx convert F --to md -o NEW.md [--notes] [--skip-hidden]`, `--to html` | `prs.to_markdown(notes=True, skip_hidden=True)`, `prs.to_html(...)` | one section per slide, in outline order |
| Embedded fonts | `rpptx font list/embed/remove/settings` | `prs.embedded_fonts`, `prs.embed_font(typeface, "ExampleSans-Regular.ttf", "regular", license=rpptx.FontEmbeddingLicense(...))`, `remove_embedded_font`, `remove_font` | rule 7; leave `embed_true_type_fonts` on; embed last (PowerPoint for Mac drops embedded fonts when it saves); a compressed embedded face is not rendered: pass its file with `font_dir` |
| From a .potx template | | `rpptx.Presentation("t.potx")`, then save as .pptx | the save writes the content type the extension names |
| Validity | `rpptx validate F` | `prs.validate()` | a tuple of issues (`kind`, `message`), empty when valid |
| What changed | `rpptx diff A B` | | slide text only |
| An element with no API | | `shape.xml` / `shape.replace_xml(xml)`, also on text frames, slides, layouts | instead of lxml; a bad replacement raises and changes nothing |

Commands and signatures: `references/cli.md`, `references/python-api.md` (and the `.pyi` stubs it points
to). Tested examples: `references/recipes.md`. Ready-made commands: `scripts/pptx_ops.py` (`replace`,
`overflow`, `shapes`). In a reference, look a name up (`grep -n NAME references/python-api.md`) rather than
reading the whole file.

## When rpptx cannot do a step

`references/gaps.md` lists each known gap with its workaround and fallback. When a step is blocked: say so in
one line to the user, use the workaround if there is one, otherwise do **that step only** with the default
`pptx` skill (python-pptx), on the file rpptx saved, keep rpptx for everything else and for the verification, and report the gap
unless it is listed.

## Reporting a bug or a missing feature

When rpptx fails, refuses, loses data, renders wrong, or lacks an API you needed, and `references/gaps.md`
does not list it: reduce it to a neutral reproduction (a short script that builds its own input, with
python-pptx if needed, and prints the wrong result next to the expected one; never the user's deck, its
text, names or paths). Ask the user before writing it to their folder; then write it as
`rdocx-findings/YYYY-MM-DD-<slug>.md` in the working folder (what, expected, observed, the script, the
pinned commit from `rdocx_env.py status`), or give it in the reply. Do not open upstream issues yourself.
