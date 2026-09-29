# rpptx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rpptx <command>`. Exit codes: 0
success; 1 error, with the message on stderr; 2 bad command line; 101 an internal panic (report it). Slide
numbers on the command line are one-based.

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/remove` (`-o` required) | refused, exit 1, "output already exists", the input included |
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

## Editing

```bash
rpptx replace F -p OLD -v NEW --expect N -o OUT
```
Literal replacement in slide text, groups, tables and speaker notes, keeping run formatting. With
`--expect N` nothing is written unless exactly N replacements were made: always pass it.

```bash
rpptx comment add F --slide N --author NAME [--initials I] --text T --date 2026-09-27T12:00:00Z -o OUT [--json]
rpptx comment reply F --id ID --author NAME --text T --date RFC3339 -o OUT
rpptx comment resolve F --id ID -o OUT
rpptx comment remove F --id ID -o OUT      # a thread with its replies, or one reply
```
Comments are anchored on the slide. The author is reused by name or added to the author list. Ids are the
GUIDs shown by `comment list --json`.

## Rendering

```bash
rpptx convert F --to pdf -o NEW.pdf
rpptx convert F --to png|jpeg|tiff -o OUT.png [--slides 1,3-5] [--dpi 150] [--quality 90] [--transparent]
rpptx render F -o NEW_DIR [--slide N] [--dpi 150] [--format png|jpeg|tiff]
rpptx thumbnail F -o NEW.png               # slide 1, 320 pixels wide
```
`convert` to images writes `OUT.png` for one slide, `OUT_001.png`, `OUT_002.png`... for several; `render`
writes `NEW_DIR/<name>_slide<N>.png`. Solid slide backgrounds are missing from the PDF but present in PNG
output (gap).
