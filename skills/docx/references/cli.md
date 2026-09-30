# rdocx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rdocx <command>`. Every command takes
one file (two for `diff` and `compare`). Exit codes: 0 success; 1 error, with the message on stderr; 2 bad
command line; 101 an internal panic (report it).

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/remove`, `revision accept/reject`, `compare`, `toc rebuild` (`-o` required) | refused, exit 1, "output already exists": they never overwrite anything, the input included |
| `render` (into `-o DIR`, default the current folder), `convert` (`-o` optional for pdf/md/html: default, the input name with the new extension) | refused the same way; `--force` replaces an existing output, never the input. Render into a new, empty folder each time |

**JSON.** `--json` output carries `"schema": 1`. The editing commands that take `--json` (all but `replace`)
print an operation record: `output`, `scope` and the result (`comment_id`, `main_story_revisions`,
`diagnostics`, ...).

## Reading

| Command | Output |
|---|---|
| `rdocx text F` | plain text, accepted view of tracked changes (insertions in, deletions out): the body first, one line per paragraph, a table row on one line, then every other story (text boxes, headers, footers, footnotes, endnotes, comments), each part under a line such as `--- header (/word/header1.xml) ---`, a part without text left out. A story part that cannot be read is left out with a warning on stderr (exit 0): `validate` names it |
| `rdocx text --json F` | `{"paragraphs": [...], "revision_view": "accepted", "scope": "all-supported-stories", "stories": [...]}`: `paragraphs` holds the body and its tables, as if tracked changes were accepted. Each paragraph: `body_index` (top-level block), `path` (list of `{kind, index}` from the block down to the paragraph, through tables and content controls; `[]` for a top-level paragraph), `style` (style id), `numbering` (`[num_id, level]` or null), `text`, `runs` (`index`, `text`, `formatting`: `bold`, `italic`, `underline`, `strike`, `size_points`, `font`, `color`, `highlight`, `language`, `style`, or null; the text inside smart tags, inline custom XML and simple fields is in `text`, not in `runs`). `stories` lists every other story: `kind` (`text_box`, `header`, `footer`, `footnote`, `endnote`, `comment`), `part_name`, `owner_index`, `items` (`index_path`, `kind`, `text`). When a story part cannot be read, `scope` is `main` and `stories` is empty |
| `rdocx inspect [--json] F` | paragraph and table counts, `content_elements`, `styles_used`, `metadata` (title, author, subject, keywords); no images |
| `rdocx layout --json F` | `body_items`: for each top-level block, `body_index`, `kind`, `fragments` (`physical_page`, `displayed_page`, `x`, `y`, `width`, `height` in points) |
| `rdocx comment list [--json] F` | `comments`: `id`, `author`, `initials`, `date`, `text`, `parent_id`, `resolved`, in package order |
| `rdocx revision list [--json] F` | `revisions`: `id`, `kind` (insertion, deletion, paragraph_property_change, run_property_change, section_property_change, ...), `author`, `timestamp`, `story` (`kind`, `part_name`, `owner_index`): every supported story |
| `rdocx diff A B` | paragraphs that differ, by one-based position, `-` and `+` lines, with the paragraph and table counts of each file |
| `rdocx validate F` | package and schema invariants: every XML part the main document relates to must be well formed (a truncated header fails, naming the part), and every paragraph, character or table style id used by the body, headers, footers, notes and comments must be defined. Warnings (empty paragraphs, missing title) keep exit 0 |

## Editing

```bash
rdocx replace F -p OLD -v NEW --expect N -o OUT
```
Literal, run-aware (a match may cross runs; the new text takes the first matched run's formatting). Reaches
the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes, headers
and footers with their tables, footnotes and endnotes, counting a header or footer text once per variant part
(default, first page, even) and a Word text box once. A match that crosses the edge of a content control, an
insertion, a simple field or a smart tag is not replaced. Always pass `--expect N`: nothing is written unless
exactly N matches were replaced (without it, any count is written). Occurrences it cannot reach are not in N:
`docx_ops.py replace` checks that none is left behind.

```bash
rdocx comment add F --start-paragraph P --start-run S --end-paragraph Q --end-run E --author A [--initials I] --text T [--date RFC3339] -o OUT [--json]
rdocx comment reply F --id ID --author A --text T [--date RFC3339] -o OUT
rdocx comment resolve F --id ID -o OUT
rdocx comment remove F --id ID -o OUT        # removes the comment and its replies
```
`P` and `Q` are **body indices** (`body_index` of `text --json`, counting tables); runs are zero-based run
boundaries, half-open: `--start-run 1 --end-run 2` covers run 1. Runs are counted as `text --json` lists them,
those inside an inline content control or a tracked insertion included. Without `--date` a comment is
undated: always pass one (`2026-09-29T12:00:00Z`). Prefer `docx_ops.py comment`, which finds the text,
splits runs, dates the comment and refuses a misplaced anchor. Comment ids stay stable across saves.

```bash
rdocx revision accept F -o OUT [--id ID] [--author NAME] [--start-date RFC3339] [--end-date RFC3339] [--json]
rdocx revision reject F -o OUT [same selectors]
```
Without a selector, every revision of every supported story (body, headers, footers, notes). Dates are
inclusive. The `--json` record's `resolved` count covers every story: the way to count revisions outside
the body.

```bash
rdocx compare ORIGINAL EDITED --author NAME --timestamp 2026-09-27T12:00:00Z -o OUT [--granularity run|word|character]
      [--ignore-comments] [--ignore-formatting] [--ignore-whitespace] [--ignore-fields] [--ignore-story KIND] [--json]
```
Writes ORIGINAL with EDITED's differences as tracked changes by NAME at the timestamp, in every supported
story. JSON record: `main_story_revisions`, `diagnostics`, `scope`, `output`. The default granularity is
`run`: one changed word shows as its whole run deleted and re-inserted. Pass `--granularity word` for a
redline a person reads. A pair whose comments differ is refused unless `--ignore-comments` (the original's
comments are kept, the edited file's dropped). A pair whose edited side had
its TOC rebuilt compares, and the rebuilt TOC entries show as revisions next to the edit: check the exit code
and read the redline before presenting it.

```bash
rdocx toc rebuild F -o OUT [--json]
```
Rebuilds existing TOC fields from the headings and rdocx's pagination (prints `Entries: N`). The entry of a numbered heading gets a left stop after its
number, so its title stays on the left when the TOC style has only the page-number stop.

## Rendering and conversion

```bash
rdocx convert F --to pdf -o NEW.pdf [--font-dir DIR]
rdocx convert F --to png|jpeg|tiff -o OUT.png [--dpi 150] [--pages 1,3-5] [--quality 90] [--transparent]
rdocx convert F --to md|html -o NEW.md
rdocx render F -o NEW_DIR [--dpi 150] [--pages 1,3-5 | --page 0] [--format png|jpeg|tiff]
```
Images from `convert`: one page goes to `OUT.png`, several to `OUT_001.png`, `OUT_002.png`... (tiff: one
multi-page file). `render` writes `NEW_DIR/<name>_page<N>.png` with N one-based; `--pages` is one-based,
`--page` zero-based. `--font-dir` adds fonts for PDF output; without it, rdocx uses its bundled
metric-compatible families (Liberation for Arial, Times New Roman, Courier New; Carlito for Calibri; Caladea
for Cambria) and the system fonts. Markdown and HTML hold the body, then text boxes, headers, footers,
footnotes and endnotes, one section per part, and leave comments out.
