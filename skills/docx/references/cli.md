# rdocx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rdocx <command>`. Every command takes
one file (two for `diff` and `compare`). Exit codes: 0 success; 1 error, with the message on stderr; 2 bad
command line; 101 an internal panic (report it).

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/remove`, `revision accept/reject`, `compare`, `toc rebuild` (`-o` required) | refused, exit 1, "output already exists": they never overwrite anything, the input included |
| `render` (into `-o DIR`, default the current folder), `convert --to png/jpeg/tiff` | refused the same way: render into a new, empty folder each time |
| `convert --to pdf/md/html` (`-o` optional: default, the input name with the new extension) | **overwritten without a word, the input included** (gap cli-convert-overwrites): `convert F --to pdf -o F` destroys F. Always give a new path |

**JSON.** `--json` output carries `"schema": 1`. The editing commands that take `--json` (all but `replace`)
print an operation record: `output`, `scope` and the result (`comment_id`, `main_story_revisions`,
`diagnostics`, ...).

## Reading

| Command | Output |
|---|---|
| `rdocx text F` | plain text of the body: one line per paragraph, a table row on one line. Leaves out body-level content controls (gap sdt-text-cli), **the text of tracked insertions** (gap text-cli-tracked-insertions), headers, footers, footnotes and text boxes (gap text-other-stories). For a document under review, use `--json` |
| `rdocx text --json F` | `{"paragraphs": [...], "revision_view": "accepted", "scope": "main"}`: the body and its tables, as if tracked changes were accepted (insertions in, deletions out). Each paragraph: `body_index` (top-level block), `path` (list of `{kind, index}` from the block down to the paragraph, through tables and content controls; `[]` for a top-level paragraph), `style` (style id), `numbering` (`[num_id, level]` or null), `text`, `runs` (`index`, `text`, `formatting`: `bold`, `italic`, `underline`, `strike`, `size_points`, `font`, `color`, `highlight`, `language`, `style`, or null). Headers, footers, footnotes and text boxes are not in it: `docx_ops.py text F` |
| `rdocx inspect [--json] F` | paragraph and table counts, `content_elements`, `styles_used`, `metadata` (title, author, subject, keywords); no images |
| `rdocx layout --json F` | `body_items`: for each top-level block, `body_index`, `kind`, `fragments` (`physical_page`, `displayed_page`, `x`, `y`, `width`, `height` in points) |
| `rdocx comment list [--json] F` | `comments`: `id`, `author`, `initials`, `date`, `text`, `parent_id`, `resolved`, in package order |
| `rdocx revision list [--json] F` | `revisions`: `id`, `kind` (insertion, deletion, paragraph_property_change, run_property_change, section_property_change, ...), `author`, `timestamp`; main story only (gap revisions-main-story) |
| `rdocx diff A B` | paragraphs that differ, by one-based position, `-` and `+` lines, with the paragraph and table counts of each file |
| `rdocx validate F` | schema and package invariants of the main parts; warnings (empty paragraphs, missing title) keep exit 0. It passes a truncated header or footer and a dangling style id (gap validate-parts): verification is `validate`, re-open in Python, render |

## Editing

```bash
rdocx replace F -p OLD -v NEW --expect N -o OUT
```
Literal, run-aware (a match may cross runs; the new text takes the first matched run's formatting). Reaches
the body, its tables and text boxes, headers and footers, counting a header or footer text once per variant
part (default, first page, even) and a Word text box twice (gap textbox-alternate-content). Does not reach
text inside content controls (gap sdt-replace), tracked insertions (gap replace-tracked-insertions),
footnotes or endnotes (gap replace-footnotes), tables in headers and footers (gap
replace-header-footer-tables), simple fields and smart tags (gap text-wrapped-runs). Always pass
`--expect N`: nothing is written unless exactly N matches were replaced (without it, any count is written).
Occurrences it cannot reach are not in N: `docx_ops.py replace` checks that none is left behind.

```bash
rdocx comment add F --start-paragraph P --start-run S --end-paragraph Q --end-run E --author A [--initials I] --text T -o OUT [--json]
rdocx comment reply F --id ID --author A --text T -o OUT
rdocx comment resolve F --id ID -o OUT
rdocx comment remove F --id ID -o OUT        # removes the comment and its replies
```
`P` and `Q` are **body indices** (`body_index` of `text --json`, counting tables); runs are zero-based run
boundaries, half-open: `--start-run 1 --end-run 2` covers run 1. **The run index skips runs inside an inline
content control or a tracked insertion, while `text --json` counts them** (gap comment-runposition-sdt): in
such a paragraph the comment lands on the wrong text without an error. The CLI has no `--date`, so its
comments are undated (gap comment-date-cli). Prefer `docx_ops.py comment`, which splits runs, dates the
comment and refuses a misplaced anchor. Comment ids stay stable across saves.

```bash
rdocx revision accept F -o OUT [--id ID] [--author NAME] [--start-date RFC3339] [--end-date RFC3339] [--json]
rdocx revision reject F -o OUT [same selectors]
```
Without a selector, every revision of every supported story (body, headers, footers, notes). Dates are
inclusive. The `--json` record's `resolved` count covers every story: the way to count revisions outside
the body.

```bash
rdocx compare ORIGINAL EDITED --author NAME --timestamp 2026-09-27T12:00:00Z -o OUT [--json]
```
Writes ORIGINAL with EDITED's differences as tracked changes by NAME at the timestamp. JSON record:
`main_story_revisions`, `diagnostics`, `scope`, `output`. Whole-run granularity and several refusals (gaps):
check the exit code and read the redline before presenting it.

```bash
rdocx toc rebuild F -o OUT [--json]
```
Rebuilds existing TOC fields from the headings and rdocx's pagination (prints `Entries: N`). Fails on a
fresh open when the TOC field runs carry `w:rsid*`, as Word writes them (gap toc-rsid-field-runs);
`docx_ops.py toc IN OUT` works around it.

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
for Cambria) and the system fonts. Markdown and HTML hold the body only, like `text`.
