# rdocx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rdocx <command>`. Every command takes
one file (two for `diff` and `compare`). Exit codes: 0 success; 1 error, with the message on stderr; 2 bad
command line; 101 an internal panic (report it).

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/move/remove`, `revision accept/reject`, `compare`, `toc rebuild`, `font embed/remove/settings` (`-o` required) | refused, exit 1, "output already exists": they never overwrite anything, the input included |
| `render` (into `-o DIR`, default the current folder), `convert` (`-o` optional for pdf/md/html and for Markdown or HTML input: default, the input name with the new extension) | refused the same way; `--force` replaces an existing output, never the input. Render into a new, empty folder each time |

**JSON.** `--json` output carries `"schema": 1`. The editing commands that take `--json` print an operation
record: `output`, `scope` and the result (`comment_id`, `main_story_revisions`, `pairs` and `total_count` for
`replace`, `diagnostics`, ...).

## Reading

| Command | Output |
|---|---|
| `rdocx text F` | plain text, accepted view of tracked changes (insertions in, deletions out; a paragraph or a table row that a tracked deletion removes prints no line): the body first, one line per paragraph, a table row on one line, then every other story (text boxes, headers, footers, footnotes, endnotes, comments), each part under a line such as `--- header (/word/header1.xml) ---`, a part without text left out. A story part that cannot be read is left out with a warning on stderr (exit 0): `validate` names it |
| `rdocx text --json F` | `{"paragraphs": [...], "revision_view": "accepted", "scope": "all-supported-stories", "stories": [...]}`: `paragraphs` holds the body and its tables, as if tracked changes were accepted (a paragraph that a tracked deletion removes is left out, the others keep their `body_index`). Each paragraph: `body_index` (top-level block), `path` (list of `{kind, index}` from the block down to the paragraph, through tables and content controls; `[]` for a top-level paragraph), `style` (style id), `numbering` (`{num_id, level}` or null), `text`, `runs` (`index`, `text`, `formatting`: `bold`, `italic`, `underline`, `strike`, `size_points`, `font`, `color`, `highlight`, `language`, `style`, or null; the text inside smart tags, inline custom XML and simple fields is in `text`, not in `runs`). `stories` lists every other story: `kind` (`text_box`, `header`, `footer`, `footnote`, `endnote`, `comment`), `part_name`, `owner_index`, `items` (`index_path`, `kind`, `text`). When a story part cannot be read, `scope` is `main` and `stories` is empty |
| `rdocx inspect [--json] F` | `paragraphs` and `tables` counts, `content_elements`, `styles_used`, `metadata` (title, author, subject, keywords), `word_count`, `character_count` (spaces included), `page_count` (rdocx's pagination), `pictures` of every story (`relationship_id`, `name`, `description` (the alt text), `title`, `decorative`, `anchored`, `width_emu`, `height_emu`, `content_type`, `story`); content controls: `doc.content_controls()` |
| `rdocx layout --json F` | `body_items`: for each top-level block, `body_index`, `kind`, `fragments` (`physical_page`, `displayed_page`, `x`, `y`, `width`, `height` in points); a paragraph that a tracked deletion removes has none |
| `rdocx comment list [--json] F` | `comments`: `id`, `author`, `initials`, `date`, `text`, `parent_id`, `resolved`, in package order; with `--json` also `anchor_text` (the accepted-view text the comment covers, paragraphs joined with a newline; null for a reply), `anchor` (`start` and `end`, each with `story_kind`, `part_name`, `owner_index`, `item_kind`, `index_path`, `direct_body_index` and `run_index`, end exclusive; null for a reply) |
| `rdocx revision list [--json] F` | `revisions`: `id`, `kind` (insertion, deletion, paragraph_property_change, run_property_change, section_property_change, ...), `author`, `timestamp`, `story` (`kind`, `part_name`, `owner_index`): every supported story |
| `rdocx diff A B [--json] [--exit-code]` | paragraphs that differ in every story, `-` and `+` lines, each labelled with its body position or its story (`[4]`, `[header default, section 1, paragraph 1]`), with the paragraph and table counts of each file. `--json`: `differences` (`change`, `story`, `location_a`, `location_b`, `text_a`, `text_b`) with `added`, `removed`, `changed` counts. Exit 0 whether or not they differ, unless `--exit-code`: 1 when they differ, as `diff` does |
| `rdocx validate F` | package and schema invariants: every XML part the main document relates to must be well formed (a truncated header fails, naming the part), and every paragraph, character or table style id used by the body, headers, footers, notes and comments must be defined; a comment with no range and no reference anywhere is an error too (`invalid comment ownership: comment N is an orphan root with no source range or reference`, exit 1, on stdout). Warnings (empty paragraphs, a paragraph holding only a picture not counted; missing title or author) keep exit 0. A paragraph with `w14:paraId` inside a block content control, as Word writes a table of contents, makes it exit 1 with "main document paragraph owners do not match validation scope" (gap docx-validate-paraid-in-content-control) |

## Editing

```bash
rdocx replace F -p OLD -v NEW --expect N -o OUT
rdocx replace F --map pairs.json -o OUT [--json]          # [{"placeholder": "{{name}}", "value": "Ada", "expect": 2}, ...]
rdocx replace F --regex -p '(\d{4})-(\d{2})-(\d{2})' -v '$3/$2/$1' --expect N -o OUT
```
Literal, run-aware (a match may cross runs; the new text takes the first matched run's formatting). Reaches
the body, its tables, content controls, tracked insertions, simple fields, smart tags and text boxes, headers
and footers with their tables, footnotes and endnotes, counting a header or footer text once per variant part
(default, first page, even) and a Word text box once. A match that crosses the edge of a content control, an
insertion, a simple field or a smart tag is not replaced. Always pass `--expect N`: nothing is written unless
exactly N matches were replaced (without it, any count is written). Occurrences it cannot reach are not in N:
`docx_ops.py replace` checks that none is left behind.
`--map` applies the pairs in order (a later pair sees what an earlier one wrote), all or nothing: a pair whose
`expect` is not met stops with an error naming it ("pair 0: expected 3 replacement(s) of ..., found 2")
and writes nothing. A pair without `expect` writes whatever it finds, nothing included: give every pair its
`expect`. `--json` gives `pairs` (each with its `count` and `expected`) and `total_count`. `--regex` reads the
placeholder as a regular expression and `$1` in the value as its first group.

```bash
rdocx fields update F -o OUT [--now 2026-10-10T09:00:00] [--json]
rdocx fill F --tag client=Acme --tag due=2026-10-31 -o OUT [--json]
rdocx meta get [--json] F
rdocx meta set F [--title T] [--author A] [--subject S] [--keywords K] [--description D] [--category C] [--custom NAME=TYPE:VALUE] [--remove-custom NAME] -o OUT [--json]
rdocx images extract F DIR [--json]
```
`fields update` recomputes DATE, TIME (at `--now`, exactly `YYYY-MM-DDTHH:MM:SS`; without it they are left
unresolved), SEQ, REF, FILENAME and the others, then the page fields from rdocx's pagination
(`update_all_fields`; `--json` lists each field's `evaluations`). `fill` sets body content controls by tag
(`--tag` repeatable; no alias: `set_content_control_value_by_alias` in Python); an unknown tag exits 1 and
writes nothing; the value is written as given, placeholder flag kept (gap
docx-content-control-value-like-word). `meta set` writes the core properties it is given and custom
properties (`--custom` repeatable, `TYPE` one of `str`, `int`, `float`, `bool`, `filetime`
(`2026-10-31T00:00:00Z`), `empty`: `Client=str:Acme`, `Count=int:4`); `meta get --json` prints them all
(`core`, `custom` with `name`, `type`, `value`). `images extract` writes each picture part once into `DIR`
(`image-000001.png`...) and lists every picture (`pictures`: `relationship_id`, `description`,
`content_type`, `width_emu`, `height_emu`, `story`, `file`); an existing file is refused, exit 1: extract into
a new folder.

```bash
rdocx comment add F --start-paragraph P --start-run S --end-paragraph Q --end-run E --author A [--initials I] --text T [--date RFC3339] -o OUT [--json]
rdocx comment add F --anchor TEXT [--occurrence K] --author A --text T [--date RFC3339] -o OUT   # K from 0, main story, cells included
rdocx comment reply F --id ID --author A --text T [--date RFC3339] -o OUT [--json]
rdocx comment resolve F --id ID -o OUT [--json]
rdocx comment move F ID --text TEXT [--occurrence K] -o OUT [--json]   # a thread root onto TEXT
rdocx comment remove F --id ID -o OUT [--json]   # removes the comment and its replies
```
`comment move` keeps the thread's id, replies and resolved state; a reply id, an unknown id or a text not
found exits 1 and writes nothing. `comment list --json` gives each thread's `anchor_text` first.
`P` and `Q` are **body indices** (`body_index` of `text --json`, counting tables); runs are zero-based run
boundaries, half-open: `--start-run 1 --end-run 2` covers run 1. Runs are counted as `text --json` lists them,
those inside an inline content control or a tracked insertion included. Without `--date` a comment is
undated: always pass one (`2026-09-29T12:00:00Z`). Prefer `docx_ops.py comment`, which finds the text,
splits runs, dates the comment and refuses a misplaced anchor. Comment ids stay stable across saves.

```bash
rdocx revision accept F -o OUT [--id ID] [--author NAME] [--start-date RFC3339] [--end-date RFC3339] [--json]
rdocx revision reject F -o OUT [same selectors] [--json]
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
redline a person reads: each changed stretch shows as one deletion then one insertion, short unchanged
words inside it folded in. A comment added or removed on the edited side replaces the paragraph it anchors on
(deleted and reinserted), a reply or a resolved state adds one insertion: accepting gives EDITED's comments,
rejecting ORIGINAL's. `--ignore-comments` keeps the original's comments, drops the edited file's, and
compares the text only. A pair whose edited side had
its TOC rebuilt compares, and the rebuilt TOC entries show as revisions next to the edit: check the exit code
and read the redline before presenting it. A chart whose data changed shows as the old chart deleted and the
new one inserted; accepting gives the edited chart, rejecting the original. `--ignore-story KIND` (repeatable:
`header`, `footer`, `comment`, `text_box`, `footnote`, `endnote`) keeps that story of ORIGINAL.

```bash
rdocx toc rebuild F -o OUT [--json]
```
Rebuilds existing TOC fields from the headings and rdocx's pagination (prints `Entries: N`). The entry of a numbered heading gets a left stop after its
number, so its title stays on the left when the TOC style has only the page-number stop.

## Fonts

```bash
rdocx font list F [--json]
rdocx font embed F --font FONT.ttf --name 'Example Sans' [--kind regular|bold|italic|bold-italic] --license-authorized \
      --license-identity 'OFL-1.1: Example Sans' [--font-key GUID] [--subsetted] -o OUT [--json]
rdocx font remove F --name 'Example Sans' [--kind bold] -o OUT [--json]
rdocx font settings F [--json]
rdocx font settings F --embed-true-type-fonts true|false|none [--embed-system-fonts true|false|none] [--save-subset-fonts true|false|none] -o OUT [--json]
```

`--font` is required for embedding. Python uses `bold_italic`, the CLI uses
`bold-italic`. Removing without `--kind` removes the whole record and its faces.
Settings reads require no output. Mutations require an output and at least one
explicit setting. `none` removes the setting, preserving its distinction from
`false`. Embedding requires explicit caller authorization and a nonempty exact
licence identity, check permission from the actual font source first. Invalid
inputs and existing outputs refuse without changing the input or publishing a
partial output.

`font list --json` returns `schema: 1` and `fonts`, it does not include settings.
Each record has `name`, `alternate_name`, `family`, `pitch` and `faces`.
Each face has `kind`, `format` (`odttf`), `part_name`, `stored_len`,
`decoded_len`, `font_key` and `subsetted`. Read the three flags separately
with `font settings --json`.

Own deterministic rendering uses supported enabled embedded faces before bundled
fallback. Explicit caller `--font-dir` faces take precedence. Word can display
only one embedded face per font name and may synthesize bold. The API and own
renderer retain all four face kinds. Verify the delivered file in Word.

## Rendering and conversion

```bash
rdocx convert F --to pdf -o NEW.pdf [--font-dir DIR] [--revision-view accepted|tracked] [--revision-style plain|legal] [--force]
rdocx convert F --to png|jpeg|tiff -o OUT.png [--dpi 150] [--pages 1,3-5] [--quality 90] [--transparent] [--revision-view tracked] [--revision-style legal] [--font-dir DIR]
rdocx convert F --to md|html -o NEW.md
rdocx convert notes.md|page.html [--to docx] [--template brand.dotx] -o NEW.docx
rdocx render F -o NEW_DIR [--dpi 150] [--pages 1,3-5 | --page 0] [--format png|jpeg|tiff] [--quality 90] [--transparent] [--revision-view tracked] [--revision-style legal] [--font-dir DIR] [--force]
rdocx layout --json F [--font-dir DIR]
```
Images from `convert`: one page goes to `OUT.png`, several to `OUT_001.png`, `OUT_002.png`... (tiff: one
multi-page file). `render` writes `NEW_DIR/<name>_page<N>.png` with N one-based; `--pages` is one-based,
`--page` zero-based. `--font-dir DIR` (`convert` to PDF or images, `render`, `layout`) puts the fonts of a folder (`.ttf`, `.otf`, `.ttc`) before every other font, an embedded face included (a file also counts as the family its file name gives, `Aileron.ttf`), a variable font drawn at the weight a run asks for (a bold run without a bold face gets a synthetic bold); a missing folder is an error, and Markdown and HTML refuse the flag. Without it, rdocx uses its bundled
metric-compatible families (Liberation for Arial, Times New Roman, Courier New; Carlito for Calibri; Caladea
for Cambria) and the system fonts. Markdown and HTML hold the body, then text boxes, headers, footers,
footnotes and endnotes, one section per part, and leave comments out.
An input named `.md` / `.markdown` is read as Markdown (CommonMark with GitHub tables), one named `.html` /
`.htm` as HTML, and written as DOCX unless `--to` names another format (a DOCX input needs `--to`). Headings
get Heading N, lists numbering, tables a header row, links and images (HTML: data URIs; Markdown: files next
to the input) are kept, block quotes get Quote and code HTML Preformatted. `--template` (Markdown input) gives
the styles, numbering, headers and footers of a DOCX or DOTX. What the import drops or approximates is a
`<location>: <message>` line on stderr (`bytes 0..23` for Markdown, `html/body/p[1]/a[2]` for HTML), exit 0:
read them.
`--quality` sets the JPEG quality (1 to 100), `--transparent` leaves unpainted PNG pixels transparent, and
`--force` replaces an existing output, never the input.
Tracked changes: every output shows the accepted view by default (insertions in, deletions out, no marks).
`--revision-view tracked` renders a PDF or images with deletions struck through, insertions underlined and a
change bar, which is how to show a redline; Markdown and HTML refuse it. `--revision-style legal` (only with
`--revision-view tracked`) colours it as legal comparison tools do: insertions blue with a double underline,
deletions red and struck through, moved text green, the change bar in the colour of the change.
