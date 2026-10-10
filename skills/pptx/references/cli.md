# rpptx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rpptx <command>`. Exit codes: 0
success; 1 error, with the message on stderr; 2 bad command line; 101 an internal panic (report it). Slide
numbers on the command line are one-based.

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/remove`, `footer`, `slide add/duplicate/remove/move/hide/show`, `notes set`, `meta set` (`-o` required) | refused, exit 1, "output already exists", the input included |
| `render` (into `-o DIR`), `convert`, `thumbnail` | refused the same way; `--force` replaces an existing output, never the input. Render into a new, empty folder each time |

## Reading

| Command | Output |
|---|---|
| `rpptx text [--notes] F` | slide text in presentation order; `--notes` adds the speaker notes after each slide |
| `rpptx text --json F` | `{"slides": [{"slide", "id", "notes", "paragraphs": [{"shape_id", "path", "level", "text", "runs": [{"index", "text", "formatting": {"bold", "italic", "underline", "color", "font", "size_points"}}]}]}]}`; `path` goes through groups, tables and paragraphs |
| `rpptx outline [--json] [--notes] F` | each slide's title and the paragraph tree of its other shapes |
| `rpptx inspect [--json] F` | layouts count, metadata (title, creator, dates), `slide_details` with, per slide, `shape_details`: every shape (id, name, kind, geometry, autofit, children for groups) |
| `rpptx comment list [--json] F` | `comments`: one flat list in slide order, threads and replies, each with `slide` (one-based), `id` (GUID), `parent_id`, `author`, `initials`, `date`, `text`, `resolved`, `status` |
| `rpptx diff A B` | slide text differences (longest common subsequence); prints nothing when the text is identical |
| `rpptx validate F` | package and PresentationML invariants |
| `rpptx meta get [--json] F` | core properties (`title`, `author`, `subject`, `keywords`, `description`, `category`, dates, `revision`...) |
| `rpptx fit [--json] F` | every text frame whose text overflows, with `needed_font_scale` (the "shrink text on overflow" scale that fits it, null below 25 %); exit 0 when all fit, 1 when one overflows, 2 on an error. Tables are not checked |

## Editing

```bash
rpptx replace F -p OLD -v NEW --expect N -o OUT [--json]
rpptx replace F --map pairs.json -o OUT [--json]   # [{"placeholder": "{{name}}", "value": "Ada", "expect": 2}, ...]
```
Literal replacement in slide text, groups, tables and speaker notes, keeping run formatting. With
`--expect N` nothing is written unless exactly N replacements were made: always pass it. `--map` applies
the pairs in order, all or nothing (the error names the failing pair); a pair without `expect` must find at
least one, `"expect": 0` allows none. `--json` prints each pair's `count`.

```bash
rpptx footer F [--slide-number] [--footer T] [--date off|auto|TEXT] [--date-format datetime1] [--skip-title] -o OUT [--json]
rpptx slide add F --layout NAME|NUMBER [--at N] -o OUT [--json]
rpptx slide duplicate/remove/hide/show F N -o OUT [--json]
rpptx slide move F N --to M -o OUT [--json]
rpptx notes set F N --text T | --from-file PATH -o OUT [--json]
rpptx meta set F [--title T] [--author A] [--subject S] [--keywords K] [--description D] [--category C] -o OUT [--json]
```
`footer` is PowerPoint's Header and Footer dialog with Apply to All: every flag is opt-in (Python's
`set_header_footer` shows the slide number by default), `--date auto` writes a date field PowerPoint refreshes,
`--skip-title` leaves title slides without them. Slide numbers are one-based; `slide add --layout` takes a
layout name or its one-based number (an unknown one lists them), `duplicate` inserts the copy after the slide,
`move N --to M` gives the final position. A `slide` edit that changes nothing still writes the output and
reports `"changed": false`. `notes set` replaces that slide's notes (one paragraph per line), creating them.

```bash
rpptx comment add F --slide N --author NAME [--initials I] --text T --date 2026-09-27T12:00:00Z -o OUT [--json]
rpptx comment reply F --id ID --author NAME --text T --date RFC3339 -o OUT [--json]
rpptx comment resolve F --id ID -o OUT [--json]
rpptx comment remove F --id ID -o OUT [--json]   # a thread with its replies, or one reply
```
Comments are anchored on the slide (Python's `slide.add_comment(..., shape_id=)` anchors one on a shape). The
author is reused by name or added to the author list. Ids are the GUIDs shown by `comment list --json`.
With `--json`, each prints an operation record: `action`, `comment_id` (and `parent_id` for a reply), `slide`,
`output`, `"schema": 1`.

## Rendering

```bash
rpptx convert F --to pdf -o NEW.pdf [--force] [--font-dir DIR]
rpptx convert F --to png|jpeg|tiff -o OUT.png [--slides 1,3-5] [--dpi 150] [--quality 90] [--transparent]
rpptx render F -o NEW_DIR [--slide N] [--dpi 150] [--format png|jpeg|tiff] [--quality 90] [--transparent] [--force] [--font-dir DIR]
rpptx thumbnail F -o NEW.png [--force] [--font-dir DIR]     # slide 1, 320 pixels wide
```
`convert` to images writes `OUT.png` for one slide, `OUT_001.png`, `OUT_002.png`... for several; `render`
writes `NEW_DIR/<name>_slide<N>.png`. `--quality` sets the JPEG quality (1 to 100), `--transparent` leaves unpainted PNG
pixels transparent. `--font-dir DIR` adds the fonts of a folder (.ttf, .otf, .ttc) before the bundled ones, to
every output; a missing folder is an error.
