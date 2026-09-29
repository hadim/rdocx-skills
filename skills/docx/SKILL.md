---
name: docx
description: "Use this skill whenever a Word file (.docx or .dotx) is involved in any way, as input or output, even when the user does not name a tool: reading or extracting text, tables, headers, footnotes, comments or tracked changes (also when the content is used elsewhere); creating a document or filling a template; find-and-replace; formatting, styles, lists, tables, pictures, hyperlinks; adding or answering comments; accepting or rejecting tracked changes; redlining two versions; tables of contents and page fields; converting to PDF, PNG, Markdown or HTML. Triggers: 'Word doc', 'word document', '.docx', 'track changes', 'redline', 'compare versions'. Runs the hash-pinned rdocx CLI and Python binding and takes precedence over the built-in docx skill and over python-docx, docx-js, lxml or LibreOffice, which it uses only for the gaps it lists; if rdocx cannot be installed in this session it says so and hands over to the built-in docx skill. Not for .doc, PDF or native Google Docs."
---

# Word documents with rdocx

rdocx reads and writes the OOXML package natively, keeps every part it does not touch byte for byte, and
renders pages itself (no LibreOffice). One pinned build serves two interfaces: the `rdocx` CLI for whole-file
operations and the `rdocx` Python module for anything finer. Use them for every .docx task. The tool is
young: the gaps that force another tool for one step are listed in `references/gaps.md`, and every new one
you meet is reported (last section).

## Setup, once per session

```bash
SKILL=/path/to/this/skill                  # the folder of this SKILL.md: the base directory shown when it loaded
python3 "$SKILL/../../scripts/rdocx_env.py" install
R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin
export PYTHONPATH="$SKILL/scripts" PYTHONDONTWRITEBYTECODE=1
$R/rdocx text --json report.docx
$R/python my_script.py                     # the Python that has rdocx and rpptx; `import docx_ops` works
```

`install` downloads the pinned build from this plugin's release when it is not already here, checks the
SHA-256 of every file against `rdocx.lock.json` (an upstream commit and its hashes) before installing, and
exits 0 once the build is installed and verified (again: "already installed").
Shell variables do not persist between commands in most agent shells: repeat the `R=` and `export` lines.
Never `pip install rdocx` from PyPI or download a binary without its hash in the lock.

**If `install` exits 2** (no verified build for this machine, for example no network access): tell the user in one line that the pinned
rdocx build is not available here, do the task with the built-in `docx` skill, and offer to build it for
next time (`install --build`, 10 to 30 minutes with a Rust toolchain: run it in the background). Details:
`../../docs/setup.md`.

## Rules that prevent damage

1. **Never write over the input.** Write to a new path, check it, then replace the original. The editing
   commands, `render` and `convert` refuse an existing output (`render` and `convert` accept `--force`, never
   for their own input); `Document.save()` replaces its target atomically but does not refuse the input:
   use `docx_ops.save_atomic(doc, out, src)`, which does. Render into a new folder each time.
2. **Every edit declares its expected count**, dry-run first, all or nothing:
   `docx_ops.replace_batch(src, out, [(old, new, n), ...])`, which also refuses when the text is still
   there after the replacement, or `rdocx replace ... --expect N`, which does not check that. rdocx's
   counts cover the body, its tables, content controls, tracked insertions, simple fields, smart tags and
   text boxes (a Word text box once), headers and footers with their tables (once per variant part),
   footnotes and endnotes.
3. **Re-fetch handles after every edit that is not a formatting setter.** Paragraphs, runs, tables and
   `story_items` all raise `StaleElementError` after an insertion, removal, clone, split, replacement,
   `paragraph.text =`, `cell.text =`, `add_*` or comment. Write `doc.paragraphs[i].runs[j]` again after
   each; never keep `p = doc.paragraphs[i]` across edits (`references/python-api.md`, "Handles").
4. **Know which index an API takes** (`references/python-api.md`, "Three index spaces"): the body index
   (`find_content_index`, `insert_paragraph`, `remove_content`, `split_run`, `RunPosition`, `rdocx comment
   add --start-paragraph`, `layout()`) counts tables and content-control blocks; `doc.paragraphs[i]` counts
   paragraphs only. `find_content_indices(text)` also returns the table of
   contents block for any heading: never delete by it.
5. **Place comments with `docx_ops.comment_on_text`**: it finds the text, splits the runs so that it forms
   whole runs, checks the anchored text on a copy first, and dates the comment. Run indices count the runs
   `Paragraph.runs` lists, those inside inline content controls and tracked insertions included. The text
   of a simple field, a smart tag or a custom XML element is in `Paragraph.text` but not in
   `Paragraph.runs`: offsets in `p.text` are not run offsets, and a comment cannot be anchored inside one.
6. **Read what reviewers see**: `rdocx text`, `rdocx text --json` or `Paragraph.text` show the accepted
   view (tracked insertions in, deletions out). Both CLI views print the body, then every other story (text
   boxes, headers, footers, footnotes, endnotes, comments), and `docx_ops.py text F` lists every paragraph
   with its part and style.
7. **Verify after writing**: `rdocx validate OUT` (every related part well formed, every style id
   defined), re-read the text you changed (`rdocx diff IN OUT`), render the touched pages to PNG and look
   at them. `Paragraph.style` takes a style id or name the document defines and raises `KeyError` for any
   other: create the style first (`doc.add_style`).
8. **The file will be opened elsewhere** (Word, Google Docs, LibreOffice): keep what you do not understand,
   edit in place rather than rebuild, and read `references/interop.md` before touching fields, content
   controls, numbering or tracked changes.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Text of the body and tables | `rdocx text --json F` | `docx_ops.textmap(F)`, `[p.text for p in doc.paragraphs]` | accepted view of tracked changes |
| Headers, footers, footnotes, text boxes, comments | `rdocx text F`, `rdocx text --json F` (`stories`), `docx_ops.py text F` | `doc.story_items`, `docx_ops.story_paragraphs(doc)` | after the body, a Word text box once |
| Where a string occurs | `docx_ops.py count F TEXT` | `docx_ops.count(doc, text)` | by story kind |
| Counts, styles used, metadata | `rdocx inspect --json F` | `doc.styles`, `doc.sections`, `doc.header_footer_variants`, `doc.core_properties` | title, author, subject, keywords, dates: read and write |
| Page of every block | `rdocx layout --json F` | `doc.layout()`, `doc.layout_page(i)` | rdocx's own pagination, close to Word's but not equal (gaps) |
| Counted replacement | `docx_ops.py replace IN OUT --edit OLD NEW N` | `docx_ops.replace_batch` | crosses runs; `rdocx replace --expect N` does not check what it could not reach |
| Regex replacement | | `doc.replace_all_regex([(pattern, repl)])` | returns the count, no contract: check it |
| Paragraph after an anchor, same format | | `doc.clone_content(doc.paragraphs[i], bi + 1)` then set run texts | copies fields, renamed bookmarks; not comment anchors |
| Plain new paragraph | | `doc.insert_paragraph(bi, text)` | Normal style |
| Delete, move a block | | `remove_content(bi)`, `pop_content(bi)` + `insert_content(bi, frag)`, `move_content(handle, bi)` | |
| Rewrite a paragraph's text | | `doc.paragraphs[i].text = text`, `doc.set_story_text(item, text)` | the setter leaves one unformatted run (paragraph style, format and comments kept), `set_story_text` keeps the first run's format |
| Format part of a run | | `docx_ops.locate` + `docx_ops.isolate`, then `.font.bold = True` | |
| Paragraph format, style, numbering | | `.paragraph_format.*`, `.style = "Heading1"` or `"Heading 1"`, `.numbering = (num_id, level)` | style id or name, checked (`KeyError`), numbering unchecked |
| New styles and lists | | `doc.add_style(name, based_on=, ...)`, `add_numbering_definition([ListLevel(...)])`, `add_numbering_instance(d)`, `link_style_to_numbering(style, num_id, level)` | |
| Tables | | `doc.add_table`, `.cell(r, c).text`, `.clone_row(i, at)`, `.remove_row(i)`, `.width` | |
| Table merges and format | | `.set_cell_grid_span(r, c, n)`, `.set_cell_vertical_merge(r, c, "restart")`, `.set_borders(...)`, `.set_column_width(c, w)`, `cell.shading = "RRGGBB"`, `row.height = w` | a span consumes empty cells only; re-fetch the table after a merge |
| Page setup | | `doc.update_section(i, margin_top=rdocx.Inches(0.5), ...)` | `doc.sections` are read-only snapshots |
| Pictures | | `doc.add_picture(bytes, name, width=, height=)`, `doc.replace_image(rid, bytes)`, `doc.set_picture_size(rid, width, height)` | `replace_image` keeps the old size: resize for a new aspect ratio |
| Hyperlinks | | `doc.hyperlinks`, `paragraph.add_hyperlink(text, url)`, `doc.set_hyperlink_url(link, url)`, `doc.remove_hyperlink(link)` | removal keeps the text, re-fetch `doc.hyperlinks` after one |
| Comments, replies, resolution | `docx_ops.py comment`; `rdocx comment list/reply/resolve/remove`, `--date` | `docx_ops.comment_on_text`, `reply_to(date=)`, `resolve_comment`, `remove_comment`; `StoryRunRange` for table cells | always date a comment |
| Tracked changes | `rdocx revision list/accept/reject --id/--author/--start-date/--end-date` | `doc.revisions`, `accept_all()`, `reject_all()`, `accept_revision_id(id)`, by author, by dates | every story: `r.story.kind` |
| Redline of two versions | `rdocx compare A B --author N --timestamp T --granularity word [--ignore-comments] -o OUT` | `a.compare(b, author, timestamp, granularity="word", ignore_comments=True)` | the default granularity replaces whole runs, differing comments are refused without the option, a rebuilt TOC is refused (gap): check the result |
| Table of contents | `docx_ops.py toc IN OUT` | `doc.rebuild_toc()` | numbered headings get their title pushed right (gap) |
| Page fields | | `doc.update_layout_backed_fields()`, `update_page_fields()` | |
| PDF | `rdocx convert F --to pdf -o NEW.pdf` | `doc.to_pdf()` | |
| PNG pages | `rdocx render F -o NEW_DIR --pages 1-3 --dpi 100` | `doc.render_pages(dpi=, pages=[0, 1])` | CLI pages one-based, Python zero-based |
| Markdown, HTML | `rdocx convert F --to md -o NEW.md` / `--to html` | | body, then text boxes, headers, footers, notes, no comments |
| Validity | `rdocx validate F` | `Document.open(F).story_items` | every related part and every style id |
| What changed between two files | `rdocx diff A B` | | by paragraph |
| New document | | `rdocx.Document()`, or a template .docx or .dotx emptied | Word's usual styles (Title, Heading 1 to 9, Quote, List Paragraph, Caption, Table Grid...), and a .dotx saved as .docx becomes a document |

Commands and full signatures: `references/cli.md`, `references/python-api.md` (and the `.pyi` stubs it
points to). Worked, tested examples for each row: `references/recipes.md`. Ready-made commands:
`scripts/docx_ops.py` (`text`, `count`, `replace`, `comment`, `toc`, `pages`).

## When rdocx cannot do a step

`references/gaps.md` lists each known gap with its workaround and its fallback. When a step is blocked:

1. Say so in one line to the user (what is blocked, which gap).
2. Use the workaround if the gap has one; otherwise do **that step only** with the built-in `docx` skill
   (python-docx, lxml or LibreOffice), on the file rdocx saved, and keep rdocx for everything else,
   including the verification.
3. Report the gap (next section) unless `references/gaps.md` already lists it.

## Reporting a bug or a missing feature

rdocx improves from what its users hit. When rdocx fails, refuses, loses data, renders wrong, or lacks an
API you needed, and `references/gaps.md` does not list it:

- Reduce it to a **neutral reproduction**: a short script that builds its own input (python-docx is fine
  for building inputs) and prints the wrong result next to the expected one. Never include the user's
  document, its text, names or paths.
- Ask the user before writing it to their folder; then write it as `rdocx-findings/YYYY-MM-DD-<slug>.md` in
  the working folder (what, expected, observed, the script, the pinned commit from `rdocx_env.py status`),
  or give it in the reply if they prefer. The findings are triaged into the plugin's tests and gap lists;
  do not open upstream issues yourself.
