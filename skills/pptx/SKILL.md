---
name: pptx
description: "Use this skill any time a PowerPoint file (.pptx or .potx) is involved in any way, as input or output, even when the user does not name a tool: reading slide text, outline, speaker notes or comments (also when the content is used elsewhere); creating a deck or filling a template; editing text without losing formatting; counted find-and-replace; moving, resizing or restyling shapes; pictures, tables, groups, connectors; adding, reordering, hiding or removing slides; comments and replies; checking that text fits its box; rendering slides to PDF or PNG. Triggers: 'PowerPoint', 'deck', 'slides', '.pptx'. Runs the hash-pinned rpptx CLI and Python binding and takes precedence over the built-in pptx skill and over python-pptx, pptxgenjs, lxml or LibreOffice, which it uses only for the gaps it lists; if rpptx cannot be installed in this session it says so and hands over to the built-in pptx skill. Not for .ppt, Keynote or native Google Slides."
---

# PowerPoint decks with rpptx

rpptx reads and writes the presentation package natively, keeps what it does not touch, lays text out with
its own line breaker and renders slides itself (no LibreOffice). One pinned build serves the `rpptx` CLI
(whole-deck operations) and the `rpptx` Python module (everything finer; its API follows python-pptx). Use
them for every .pptx task. The tool is young: the gaps that force another tool for one step are in
`references/gaps.md`, and every new one you meet is reported (last section).

## Setup, once per session

```bash
SKILL=/path/to/this/skill                  # the folder of this SKILL.md: the base directory shown when it loaded
python3 "$SKILL/../../scripts/rdocx_env.py" install
R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin
export PYTHONPATH="$SKILL/scripts" PYTHONDONTWRITEBYTECODE=1
$R/rpptx outline deck.pptx
$R/python my_script.py                     # the Python that has rpptx and rdocx; `import pptx_ops` works
```

`install` downloads the pinned build from this plugin's release when it is not already here, checks the
SHA-256 of every file against `rdocx.lock.json` before installing, and exits 0 once the build is installed
and verified. Repeat the `R=` and `export` lines in each new shell. Never `pip
install rpptx` from PyPI or download a binary without its hash in the lock.

**If `install` exits 2** (no verified build for this machine, for example no network access): tell the user in one line that the pinned
rpptx build is not available here, do the task with the built-in `pptx` skill, and offer to build it for
next time (`install --build`, 10 to 30 minutes with a Rust toolchain: run it in the background). Details:
`../../docs/setup.md`.

## Rules that prevent damage

1. **Never write over the input.** The editing commands, `render` and image `convert` refuse an existing
   output; **`rpptx convert --to pdf` and `rpptx thumbnail` overwrite anything, the input included**;
   `Presentation.save()` writes in place: use `pptx_ops.save_atomic(prs, out, src)`. Render into a new
   folder each time.
2. **Every text replacement declares its expected count**: `rpptx replace ... --expect N` (slides and
   notes), or `pptx_ops.replace_batch(src, out, [(old, new, n), ...])`, all or nothing.
3. **Re-fetch handles after every change that is not a geometry, font, paragraph or frame setter**: write
   `prs.slides[i].shapes[j]` again. Any `add_*` invalidates the handles of every slide; setting any text
   (`run.text`, `text_frame.text`, `shape.text`, `notes_text`) invalidates everything, the shape an `add_*`
   just returned included: `box = add_textbox(...); box.text_frame.text = "x"; box.left` raises.
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
| Counted replacement | `rpptx replace F -p OLD -v NEW --expect N -o OUT` | `pptx_ops.replace_batch` | keeps run formatting; no Python API (gap) |
| Edit a run, paragraph, text frame | | `run.text`, `run.font.*`, `paragraph.alignment/level/space_*/line_spacing`, `text_frame.margin_*/word_wrap/auto_size/vertical_anchor` | `shape.text = ...` drops run formatting, as in python-pptx |
| Move, resize, rotate | | `shape.left/top/width/height/rotation` | placeholders that inherit their geometry read None (gap) |
| Fill, line | | `shape.fill.solid()`, `.fill.fore_color.rgb = RGBColor(...)`, `.line.width`, `.line.color.rgb` | |
| Add shapes | | `shapes.add_textbox`, `add_shape(MSO_SHAPE.X, ...)`, `add_connector`, `add_picture`, `add_table` | new groups cannot be filled (gap) |
| Pictures | | `shape.replace_image(file)`, `shape.image.blob` | keeps position, size and crop |
| Tables | | `shape.table.cell(r, c).text`, `table.columns[k].width` | no row or column insertion, no merge from Python (gap) |
| Slides | | `slides.add_slide(layout)`, `slides.move(i, j)`, `slides.remove(slide)`, `slide.hidden` | no duplication from Python (gap) |
| Speaker notes | `rpptx text --notes F` | `slide.notes_text` (get and set) | None when the slide has no notes |
| Hyperlinks | | | no API (gap pptx-hyperlinks) |
| Metadata | `rpptx inspect --json F` | | read-only; no core properties API |
| Comments | `rpptx comment list/add/reply/resolve/remove` | `prs.add_comment_author`, `slide.add_comment`, `reply_to_comment`, `move_comment` | resolve and remove: CLI only (gap) |
| Text fit | | `prs.text_layout(width_factor=1.0)`, `pptx_ops.overflowing(F)` | |
| PDF | `rpptx convert F --to pdf -o NEW.pdf` | `prs.to_pdf()`, `prs.to_notes_pdf()` | slide backgrounds missing in the PDF (gap) |
| PNG | `rpptx render F -o NEW_DIR --slide N --dpi 100`, `rpptx convert F --to png --slides 1-3 -o NEW.png` | `prs.render_slide_to_png(i, dpi)`, `render_all_slides(dpi)` | CLI slides one-based, Python zero-based |
| From a .potx template | | `rpptx.Presentation("t.potx")`, save, then `pptx_ops.fix_template_content_type(out)` | rpptx keeps the template content type (gap) |
| Validity | `rpptx validate F` | | |
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
