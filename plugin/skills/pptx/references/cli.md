# rpptx CLI reference (pinned build)

`R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, then `$R/rpptx <command>`. Exit codes: 0
success; 1 error, with the message on stderr; 2 bad command line; 101 an internal panic (report it). Slide
numbers on the command line are one-based.

**Outputs.**

| Commands | An existing output |
|---|---|
| `replace`, `comment add/reply/resolve/remove`, `slide add/duplicate/remove/move/hide/show/import`, `notes set`, `meta set`, `font embed/remove/settings`, `sign` (`-o` required) | refused, exit 1, "output already exists", the input included |
| `render` (into `-o DIR`), `convert`, `thumbnail`, `footer` | refused the same way; `--force` replaces an existing output, never the input. Render into a new, empty folder each time |

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
| `rpptx fit [--json] [--font-dir DIR] F` | every text owner (slides, layouts, masters, notes, table cells), one line each with `overflow` and `needed_scale` (the "shrink text on overflow" scale that fits it, in hundred-thousandths: 92500 is 92.5 %); `--json`: `overflow` (true when one overflows) and `frames`, each with `owner_kind` (`slide`, `notes-slide`, `layout`, `master`...), `slide` (one-based, null for a layout or master), `shape_id`, `name`, `cell`, `overflow`, `needed_scale`. Exit 0 when all fit, 1 when one overflows, 2 on an error |

## Editing

```bash
rpptx replace F -p OLD -v NEW --expect N -o OUT [--json]
rpptx replace F --map pairs.json -o OUT [--json]   # [{"placeholder": "{{name}}", "value": "Ada", "expect": 2}, ...]
```
Literal replacement in slide text, groups, tables and speaker notes, keeping run formatting. With
`--expect N` nothing is written unless exactly N replacements were made: always pass it. `--map` applies
the pairs in order, all or nothing (the error names the failing pair); a pair
without `expect` must find at least one ("pair N: no replacements found"), `"expect": 0` allows none.
`--json` prints each pair's `count` and the `total_count`.

```bash
rpptx footer F [--slide-number] [--footer T] [--date off|auto|TEXT] [--date-format datetime1] [--date-cache TEXT] [--skip-title] -o OUT [--json]
rpptx slide add F --layout NAME|NUMBER [--at N] -o OUT [--json]
rpptx slide duplicate/remove/hide/show F N -o OUT [--json]
rpptx slide move F N --to M -o OUT [--json]
rpptx notes set F N --text T | --from-file PATH -o OUT [--json]
rpptx meta set F [--title T] [--author A] [--subject S] [--keywords K] [--description D] [--category C] -o OUT [--json]
```
`footer` is PowerPoint's Header and Footer dialog with Apply to All: every flag is opt-in (Python's
`set_header_footer` shows the slide number by default), `--date auto` writes a date field PowerPoint fills
(empty until then; `--date-cache TEXT` gives its text), `--skip-title` leaves title slides without them; `--json`
prints the number of slides, not their fields (read them with Python's `slide.header_footer`). Slide numbers are one-based; `slide add --layout` takes a
layout name or its one-based number (an unknown one lists them), `duplicate` inserts the copy after the slide,
`move N --to M` gives the final position. A `slide` edit that changes nothing still writes the output and
reports `"changed": false`. `notes set` replaces that slide's notes (one paragraph per line), creating them; the
text is written as given, a final newline included (an empty last paragraph).

```bash
rpptx slide import DEST SOURCE [--slides 1,3-5] [--at N] [--keep-source-formatting] -o OUT [--json]
```
Copies the slides of SOURCE (all, or the one-based `--slides` selection, in the order given) into DEST, appended
or from the one-based position `--at`, and writes OUT. Charts, embedded objects, comments and notes travel with
them; a slide jump whose target is not imported stops working and is reported as a diagnostic line (and in
`diagnostics` of the `--json` record, which also gives the `selected_slides` and the `inserted_slides` they
became): read it. `--keep-source-formatting` keeps the source layout, master and theme of each slide (otherwise
the copy takes DEST's layout of the same name). Sections: Python's `import_slides(..., keep_sections=True)`.

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

## Fonts

```bash
rpptx font list F [--json]
rpptx font embed F --font FONT.ttf --name 'Example Sans' [--kind regular|bold|italic|bold-italic] --license-authorized \
      --license-identity 'OFL-1.1: Example Sans' -o OUT [--json]
rpptx font remove F --name 'Example Sans' [--kind bold] -o OUT [--json]
rpptx font settings F [--json]
rpptx font settings F --embed-true-type-fonts true|false|none [--save-subset-fonts true|false|none] -o OUT [--json]
```

`--font` is required for embedding. Python uses `bold_italic`, the CLI uses
`bold-italic`. Removing without `--kind` removes the whole record and its faces;
removing the last face with `--kind` leaves an empty record. A name (or kind) that is not
embedded exits 1 and writes nothing (`rdocx font remove` prints `Removed: false` and exits 0).
Settings reads require no output. Mutations require an output and at least one
explicit setting. `none` removes the setting, preserving its distinction from
`false`. Embedding requires explicit caller authorization and a nonempty exact
licence identity, check permission from the actual font source first. Invalid
inputs and existing outputs refuse without changing the input or publishing a
partial output.

`font list --json` returns `schema: 1` and `fonts`, it does not include settings.
Each record has `name`, `panose`, `pitch_family`, `charset` and `faces`.
Each face has `kind`, `relationship_id`, `part_name`, `format`, `stored_len`
and `decoded_len` (null when undecoded). CLI format `compressed-eot` differs
from Python `compressed_eot`. Unsupported embedded payloads are preserved,
not decoded for rendering. Read the two flags with `font settings --json`.

Own deterministic rendering uses supported enabled embedded faces before bundled
fallback. Explicit caller `--font-dir` faces take precedence. Native PowerPoint
checks retain uncompressed EOT and use genuine regular and bold faces. Verify
font retention after resaving in the receiving application.

## Signatures

```bash
openssl req -x509 -newkey rsa:2048 -nodes -keyout key.pem -out cert.pem -days 365 -subj "/CN=Signer"   # a test pair
rpptx sign F.pptx --key key.pem --cert cert.pem -o SIGNED.pptx [--json]
rpptx verify-signatures [--json] F.pptx
```
`sign` writes a package signature over every part and relationship (`_xmlsignatures/sig1.xml`), with a PKCS#8 RSA
key and an X.509 certificate, DER or one unencrypted PEM block each; it refuses an existing output, a package
already signed and one with external relationships. `verify-signatures` prints `status` (`valid`, `invalid`,
`partial`, `unsigned`) and `signatures` (`signer` with `subject`, `issuer`, `serial_number`;
`cryptographically_valid`, `coverage_complete`, `covered_parts`, `issues` with `kind` such as `changed_reference`);
exit 0 only when at least one signature is valid and complete and none is invalid or partial, 1 otherwise (an
unsigned file included). Certificate trust (the chain, revocation) is not checked: say so when you report a
signature as valid. Sign last: any later edit, by rdocx or by an application, invalidates it. In Python:
`prs.sign(key_bytes, cert_bytes)` then `save`, `prs.verify_signatures()` (`references/python-api.md`).

## Rendering

```bash
rpptx convert F --to pdf -o NEW.pdf [--force] [--font-dir DIR]
rpptx convert F --to md|html -o NEW.md [--notes] [--skip-hidden] [--embed-images=false] [--force]
rpptx convert F --to png|jpeg|tiff -o OUT.png [--slides 1,3-5] [--dpi 150] [--quality 90] [--transparent]
rpptx render F -o NEW_DIR [--slide N] [--dpi 150] [--format png|jpeg|tiff] [--quality 90] [--transparent] [--force] [--font-dir DIR]
rpptx thumbnail F -o NEW.png [--force] [--font-dir DIR]     # slide 1, 320 pixels wide
```
`convert` to images writes `OUT.png` for one slide, `OUT_001.png`, `OUT_002.png`... for several; `render`
writes `NEW_DIR/<name>_slide<N>.png`. `--quality` sets the JPEG quality (1 to 100), `--transparent` leaves unpainted PNG
pixels transparent. `--font-dir DIR` adds the fonts of a folder (.ttf, .otf, .ttc) before the bundled ones, to
every output; a missing folder is an error.
`--to md` writes CommonMark with GitHub tables, one section per slide headed by its title (an anchor before
each), the rest in the order `outline` prints: paragraphs as list items nested by level with their bold, italic,
strikethrough and external links, tables with the first row as header, pictures as images with their alt text.
`--to html` writes the same as one HTML document, a `section` per slide. Pictures are embedded as data URIs;
`--embed-images=false` writes them as files next to the output instead. `--notes` adds the speaker notes (a
"Notes" heading and a block quote, an `aside`), `--skip-hidden` leaves hidden slides out; the three are refused
for PDF and images. The Markdown keeps whitespace exact (the spaces of titles, formatted runs and notes as
`&#32;`, emphasis as `<strong>` / `<em>`): to read a deck, `text` or `outline` is plainer.
