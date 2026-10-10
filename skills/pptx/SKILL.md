---
name: pptx
description: "Use this skill any time a PowerPoint file (.pptx or .potx) is involved in any way, as input or output, even when the user does not name a tool: reading slide text, outline, speaker notes or comments (also when the content is used elsewhere); creating a deck or filling a template; editing text without losing formatting; counted find-and-replace; moving, resizing or restyling shapes; pictures, tables, groups, connectors; adding, reordering, hiding or removing slides; comments and replies; checking that text fits its box; rendering slides to PDF or PNG. Triggers: 'PowerPoint', 'deck', 'slides', '.pptx'. Runs the hash-pinned rpptx CLI and Python binding and takes precedence over the built-in pptx skill and over python-pptx, pptxgenjs, lxml or LibreOffice, which it uses only for the gaps it lists; if rpptx cannot be installed in this session it says so and hands over to the built-in pptx skill. Not for .ppt, Keynote or native Google Slides."
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

`install` downloads the pinned build from rdocx's releases when it is not already here, checks the
SHA-256 of every file against `rdocx.lock.json` before installing, and exits 0 once the build is installed
and verified. Repeat the `R=` and `export` lines in each new shell. Never `pip
install rpptx` from PyPI or download a binary without its hash in the lock.
**Windows**: run these lines in Git Bash, with `python` (or `py -3`) where they say `python3`; the quotes
around the skill folder keep its backslashes. The CLIs are `rdocx.exe` and `rpptx.exe`, which `$R/rdocx` and
`$R/rpptx` reach in Git Bash.

**If `install` exits 2** (no verified build for this machine, for example no network access): tell the user in one line that the pinned
rpptx build is not available here, do the task with the built-in `pptx` skill, and offer to build it for
next time (`install --build`, 10 to 30 minutes with a Rust toolchain: run it in the background). Details:
`../../docs/setup.md`.

## Rules that prevent damage

1. **Never write over the input.** The editing commands, `render`, `convert` and `thumbnail` refuse an
   existing output (`convert`, `render` and `thumbnail` accept `--force`, never for their own input);
   `Presentation.save()` does not refuse the input: use `pptx_ops.save_atomic(prs, out, src)`. Render into a
   new folder each time.
2. **Every text replacement declares its expected count**: `rpptx replace ... --expect N` (slides and
   notes), `prs.try_replace_text(old, new, expect=n)`, or `pptx_ops.replace_batch(src, out, [(old, new, n),
   ...])`, all or nothing.
3. **Re-fetch the handles an edit retires**: each edit retires the handles of the kind it renumbers and
   those below it (slide > shape > paragraph > run), in every slide. Appends (`add_*`, `add_slide`) and
   setters retire none; `text_frame.text =` / `shape.text =` retire paragraphs and runs, a replacement
   retires runs; removing, moving or grouping shapes retires shapes; moving or removing slides retires
   everything. A retired handle raises `StaleElementError` naming the call: write `prs.slides[i].shapes[j]`
   again.
4. **Walk groups**: `slide.shapes` lists top-level shapes; a group's children are in `shape.shapes`
   (`shape_type == MSO_SHAPE_TYPE.GROUP`, value 6). `pptx_ops.walk(slide.shapes)` yields all of them.
5. **Check the fit after any text change**: `prs.text_layout()` (or `pptx_ops.py overflow`) reports every
   frame whose text overflows, with rpptx's own line breaks; then render the slide and look at it.
6. **The deck will be opened elsewhere** (PowerPoint, Google Slides, Keynote, LibreOffice): edit in place,
   keep placeholders and layouts, never rebuild a deck to change it.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Slide text | `rpptx text [--notes] F` | `shape.text`, `shape.text_frame.paragraphs[k].runs` | |
| Text with structure and formatting | `rpptx text --json F` | | paragraphs with `path`, `shape_id`, `level`, runs with bold, italic, underline, colour, font, size; notes |
| Titles and outline | `rpptx outline [--json] [--notes] F` | `slide.shapes.title` | |
| Structure, shapes, metadata | `rpptx inspect --json F` | `pptx_ops.py shapes F [--slide N]` | shape tree with ids, names, geometry |
| Counted replacement | `rpptx replace F -p OLD -v NEW --expect N -o OUT` | `prs.try_replace_text(old, new, expect=n)`, `pptx_ops.replace_batch` | keeps run formatting |
| Replacement in one slide or one frame | | `slide.try_replace_text(old, new, expect=n)`, `shape.text_frame.try_replace_text(old, new, expect=n)` | the slide's notes too unless `notes=False`; same all-or-nothing count |
| Edit a run, paragraph, text frame | | `run.text`, `run.font.*`, `paragraph.alignment/level/space_*/line_spacing`, `text_frame.margin_*/word_wrap/auto_size/vertical_anchor` | `shape.text = ...` drops run formatting, as in python-pptx |
| Move, resize, rotate | | `shape.left/top/width/height/rotation` | a placeholder that inherits its geometry reads None: `shape.effective_geometry()` gives it, and a setter copies it first |
| Fill, line | | `shape.fill.solid()`, `.fill.fore_color.rgb = RGBColor(...)`, `.line.width`, `.line.color.rgb` | |
| Dashes, arrowheads | | `.line.dash_style = MSO_LINE_DASH_STYLE.DASH`, `.line.tail_end.type = MSO_ARROWHEAD_STYLE.TRIANGLE`, `.width`, `.length`; `head_end` likewise | enums in `rpptx.enum.dml` |
| Shadow | | `shape.shadow.visible = True`, `.color.rgb`, `.alpha`, `.blur_radius`, `.distance`, `.direction`, `.align` | writes `a:outerShdw`; `shadow.inherit = False` removes the theme's shadow |
| Connector without the theme effect | | `connector.theme_effect_index = 0` | `add_connector` references the theme's effect 1, an outer shadow in the default theme |
| Change a shape's preset | | `shape.auto_shape_type = MSO_SHAPE.RECTANGLE` | autoshapes only |
| Add shapes | | `shapes.add_textbox`, `add_shape(MSO_SHAPE.X, ...)`, `add_connector`, `add_picture`, `add_table`, `add_group_shape()` | a group's `shapes` take the same `add_*` calls, re-fetch the group after each; `add_shape` writes python-pptx's theme style (accent1 fill and line, theme effect 2), `add_textbox` none |
| Z-order | | `shapes.move(from_, to)` | index 0 is the back |
| Pictures | | `shape.replace_image(file)`, `shape.image.blob` | keeps position, size and crop |
| Tables | | `shape.table.cell(r, c).text`, `cell.text_frame.paragraphs[k].runs[j].font`, `cell.vertical_anchor`, `.merge(other)`, `.fill`, `table.columns[k].width`, `table.rows[k].height`, `table.rows.add_row(i)`, `rows.remove(row)`, `table.columns.add_column(i)`, `columns.remove(col)` | a new row or column copies a neighbour's size, re-fetch the table after each |
| Slides | | `slides.add_slide(layout)`, `slides.duplicate(slide)`, `slides.move(i, j)`, `slides.remove(slide)`, `slide.hidden` | |
| Import a slide from another deck | | `prs.slides.import_slide(other.slides[k], layout=prs.slide_layouts[j], index=None)` | without `layout=`, a layout of the same name must exist here, else `RpptxError` |
| Speaker notes | `rpptx text --notes F` | `slide.notes_text` (get and set) | None when the slide has no notes |
| Hyperlinks, slide jumps | | `run.hyperlink.address`; `shape.click_action.hyperlink.address = url`, `shape.click_action.target_slide = prs.slides[k]` | any shape, group members included; set None to clear |
| Metadata | `rpptx inspect --json F` | | read-only; no core properties API |
| Comments | `rpptx comment list/add/reply/resolve/remove` | `prs.add_comment_author`, `slide.add_comment`, `reply_to_comment`, `resolve_comment`, `remove_comment`, `move_comment` | `slide.add_comment(..., shape_id=sh.shape_id)` anchors on a shape (the CLI on the slide); add the author first |
| Text fit | | `prs.text_layout(width_factor=1.0)`, `pptx_ops.overflowing(F)` | rpptx's line breaks; `width_factor=0.95` for a margin |
| PDF | `rpptx convert F --to pdf -o NEW.pdf [--font-dir DIR]` | `prs.to_pdf(font_dir=)`, `prs.to_notes_pdf()` | `--font-dir` / `font_dir=` give fonts that are not installed, to every output and to the fit check |
| PNG | `rpptx render F -o NEW_DIR --slide N --dpi 100`, `rpptx convert F --to png --slides 1-3 -o NEW.png` | `prs.render_slide_to_png(i, dpi)`, `render_all_slides(dpi)` | CLI slides one-based, Python zero-based |
| From a .potx template | | `rpptx.Presentation("t.potx")`, then save as .pptx | the save writes the content type the extension names |
| Validity | `rpptx validate F` | `prs.validate()` | a tuple of issues (`kind`, `message`), empty when valid |
| What changed | `rpptx diff A B` | | slide text only |

Commands and signatures: `references/cli.md`, `references/python-api.md` (and the `.pyi` stubs it points
to). Tested examples: `references/recipes.md`. Ready-made commands: `scripts/pptx_ops.py` (`replace`,
`overflow`, `shapes`).

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
