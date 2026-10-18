---
name: docx
description: "Use this skill whenever a Word file (.docx or .dotx) is involved, as input or output, even when the user names no tool: reading its text, tables, comments or tracked changes, also to use them elsewhere; creating a document (also from Markdown or HTML) or filling a template; find-and-replace; formatting, styles, tables, pictures, charts, comments; accepting or rejecting tracked changes; redlining two versions; tables of contents and page fields; converting to PDF, PNG, Markdown or HTML. Triggers: 'Word doc', '.docx', 'track changes', 'redline', 'compare versions'. Uses the rdocx CLI and Python binding and takes precedence over the built-in docx skill, python-docx, docx-js and LibreOffice. Not for .doc, PDF or Google Docs."
---

# Word documents with rdocx

rdocx reads and writes the OOXML package natively, keeps every part it does not touch byte for byte, and
renders pages itself (no LibreOffice). One pinned build serves two interfaces: the `rdocx` CLI for whole-file
operations and the `rdocx` Python module for anything finer. Use them for every .docx task. Most python-docx
code runs as is; the rest raises an error that names the rdocx way (`references/python-api.md`,
"python-docx names"). The tool is young: the gaps that force another tool for one
step are in `references/gaps.md`, and every new one you meet is reported (last section).

## Setup, once per session

```bash
SKILL='/path/to/this/skill'                # the folder of this SKILL.md: the base directory shown when it loaded
python3 "$SKILL/../../scripts/rdocx_env.py" install
R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin
export PYTHONPATH="$SKILL/scripts" PYTHONDONTWRITEBYTECODE=1
$R/rdocx text --json report.docx
$R/python my_script.py                     # the Python that has rdocx and rpptx; `import docx_ops` works
```

`install` downloads the pinned build when missing, checks every file's SHA-256 against `rdocx.lock.json`
and exits 0 once verified. Shell variables do not persist between
commands in most agent shells: repeat the `R=` and `export` lines. Never `pip install rdocx` from PyPI or
download a binary without its hash in the lock. **Windows**: Git Bash, with `python` (or `py -3`) for
`python3`; `$R/rdocx` reaches `rdocx.exe` (`../../docs/setup.md`).

**If `install` exits 2** (no verified build for this machine, for example no network access): tell the user in
one line that the pinned rdocx build is not available here, do the task with the built-in `docx` skill, and
offer to build it for next time (`install --build`, 10 to 30 minutes with a Rust toolchain, in the
background).

## Rules that prevent damage

1. **Never write over the input.** Write to a new path, check it, then replace the original. The editing
   commands, `render` and `convert` refuse an existing output (`--force` for `render` and `convert`, never
   for their input); `Document.save()` does not refuse the input: use `docx_ops.save_atomic(doc, out,
   src)`, which does. Render into a new folder each time.
2. **Every edit declares its expected count**, dry-run first, all or nothing:
   `docx_ops.replace_batch(src, out, [(old, new, n), ...])`, which also refuses when the text is still
   there after the replacement, or `rdocx replace ... --expect N`, which does not check that (what each one
   reaches: `references/cli.md`, Editing).
3. **Re-fetch handles after an edit.** Formatting setters and page setup keep them; appends (`add_paragraph`,
   `add_run`, `p.text = ...`) retire every `Run` and story item; inserts, removals, row and column edits,
   replacements, pictures, comments and font edits retire every handle: the next use raises
   `StaleElementError`. Re-navigate from `doc` (`doc.paragraphs[i].runs[j]`, `doc.section(i).header`;
   `references/python-api.md`, "Handles").
4. **Know which index an API takes** (`references/python-api.md`, "Three index spaces"): the body index
   (`find_content_index`, `insert_paragraph`, `remove_content`, `split_run`, `RunPosition`, `rdocx comment
   add --start-paragraph`, `layout()`) counts tables and content-control blocks; `doc.paragraphs[i]` counts
   paragraphs only. `find_content_indices(text)` also returns the table of contents block for any heading:
   never delete by it.
5. **Place comments with `docx_ops.comment_on_text`**: it finds the text, splits the runs so that it forms
   whole runs, checks the anchored text on a copy first, and dates the comment. Offsets in `p.text` are not run
   offsets (the text of simple fields, smart tags and custom XML is in `Paragraph.text`, not in
   `Paragraph.runs`). For text inside a table cell pass `in_tables=True` (`--in-tables`).
6. **Read what reviewers see**: `rdocx text`, `rdocx text --json` and `Paragraph.text` show the accepted view
   (tracked insertions in, deletions out), the body then every other story; `docx_ops.py text F` lists every
   paragraph with its part and style.
7. **Verify after writing**: `rdocx validate OUT`, re-read the text you changed (`rdocx diff IN OUT`), render the
   touched pages to PNG and look at them.
8. **The file will be opened elsewhere** (Word, Google Docs, LibreOffice): keep what you do not understand,
   edit in place rather than rebuild, and read `references/interop.md` before touching fields, content
   controls, numbering or tracked changes.
9. **Build lengths with `rdocx.Pt(...)`, `Inches(...)`, `Twips(...)`**: a bare int is EMU, so `font.size = 12`,
   `space_after = 6` and a margin of `1440` raise; a line spacing of `2` (or `2.0`) is a multiple, a length is exact.
10. **Embed a font only when its licence allows embedding.** `rdocx.FontEmbeddingLicense(authorized=True,
   identity="OFL-1.1: Example Sans")` is your statement that it does: rdocx checks nothing, not even `fsType`.
   OFL, Apache 2.0 and CC0 allow it; a commercial licence often does not (`fsType` bit 0x0002: Restricted
   License). When it does not, or you cannot tell, do not embed and tell the user.

## What to use for what

| Task | CLI | Python | Notes |
|---|---|---|---|
| Text of the body and tables | `rdocx text --json F` | `docx_ops.textmap(F)`, `[p.text for p in doc.paragraphs]` | accepted view of tracked changes |
| Headers, footers, footnotes, text boxes, comments | `rdocx text F`, `rdocx text --json F` (`stories`), `docx_ops.py text F` | `doc.story_items`, `docx_ops.story_paragraphs(doc)` | after the body; a Word text box once |
| Where a string occurs | `docx_ops.py count F TEXT` | `docx_ops.count(doc, text)` | by story kind |
| Counts, styles used, metadata | `rdocx inspect --json F`, `rdocx meta get/set F` | `doc.word_count()`, `doc.page_count()`, `doc.styles`, `doc.core_properties`, `doc.custom_properties` | |
| Page of every block | `rdocx layout --json F` | `doc.layout()` | rdocx's pagination: confirm in Word |
| Counted replacement | `docx_ops.py replace IN OUT --edit OLD NEW N` | `docx_ops.replace_batch` | crosses runs; `rdocx replace --expect N` does not check what it could not reach |
| Fill a template or a form | `rdocx replace F --map pairs.json -o OUT`, `rdocx fill F --values v.json -o OUT` | `doc.render_template(data)`, `doc.bind_content_controls({tag: v})`, legacy fields | `--map`: pairs in order, each with its `expect`; controls: an unknown tag raises `KeyError` |
| Replacement in one paragraph, cell or story item | | `p.replace_text(old, new, expect=1)`, `cell.replace_text(old, new, expect=1)`, `doc.replace_text_at(item, old, new, expect=1)` | a wrong count raises, nothing changed |
| Regex replacement | `rdocx replace F --regex -p PATTERN -v '$1' --expect N -o OUT` | `doc.replace_all_regex_expected([(pattern, repl, n)])` | `$1` groups; a wrong count raises |
| Paragraph after an anchor, same format | | `doc.clone_content(p, bi + 1)`, then set run texts | not comment anchors |
| New paragraphs, headings, lists | | `doc.add_heading(text, 1)`, `doc.add_paragraph(text, style="Quote")`, `doc.add_bullet_list_item(text)`, `doc.add_numbered_list_item(text)`, `doc.insert_paragraph(bi, text)` | a list restart: a new `add_numbering_instance` |
| Delete, move a block | | `remove_content(bi)`, `pop_content(bi)` + `insert_content(bi, frag)`, `move_content(source, destination)` | cutting a comment in part raises: move the thread first |
| Rewrite a paragraph's text | | `doc.paragraphs[i].text = text`, `doc.set_story_text(item, text)` | the setter leaves one plain run; `set_story_text` keeps the first run's format |
| Format part of a run | | `docx_ops.locate` + `docx_ops.isolate`, then `runs[k].font.bold = True` | |
| Paragraph format, style, numbering | | `.paragraph_format.*`, `run.font.*`, `.style = "Heading 1"`, `.numbering = (num_id, level)` | style id or name, checked (`KeyError`: `add_style` first) |
| New or changed styles, lists | | `doc.add_style(name, based_on=, ...)`, `doc.set_style(style, bold=, ...)`, `add_numbering_definition([ListLevel(...)])`, `add_numbering_instance(d)`, `link_style_to_numbering(style, num_id, level)` | `set_style` keeps what it is not given |
| Tables | | `t = doc.add_table(r, c, style="Table Grid")`, `.cell(r, c).text`, `.add_row().cells`, `.insert_column(i, before=False)`, `.delete_column(i)`, `.remove_row(i)`, `cell.add_table(r, c)` | |
| Table merges and format | | `.set_cell_grid_span(r, c, n)`, `.set_cell_vertical_merge(r, c, "restart")`, `.set_borders(...)`, `cell.shading = "RRGGBB"` | re-fetch the table after a merge |
| Page setup, sections | | `doc.update_section(i, margin_top=rdocx.Inches(0.5), orientation="landscape")`, `doc.section(i).left_margin = rdocx.Inches(1)`, `doc.insert_section(i)` | `doc.sections` are read-only snapshots |
| Pictures | `rdocx images extract F DIR` | `doc.add_picture(image, name, w, h, options=rdocx.PictureOptions(...))`, `run.add_picture(...)`, `doc.pictures()`, `doc.replace_image(rid, bytes)` | both sizes or neither; `replace_image` keeps the old size |
| Hyperlinks | | `doc.hyperlinks`, `p.add_hyperlink(text, url)`, `p.add_internal_hyperlink(text, bookmark)`, `doc.add_heading_hyperlink(p, heading, text)`, `doc.remove_hyperlink(hyperlink)` | re-fetch `doc.hyperlinks` after a removal |
| Headers, footers, page numbers, notes | | `doc.section(i).header.add_paragraph(t)`, `run.add_field("PAGE")`, `doc.add_footnote(p, text)` | a new header holds no paragraph; `doc.remove_footnote(id)` removes a note after later edits |
| Comments, replies, resolution | `docx_ops.py comment`; `rdocx comment list/reply/resolve/remove`, `--date` | `docx_ops.comment_on_text`, `reply_to(date=)`, `resolve_comment`, `remove_comment` | always date a comment |
| What each comment points at, moving a thread | `rdocx comment list --json F` (`anchor_text`, `anchor`), `rdocx comment move F ID --text TEXT -o OUT` | `c.anchor_text`, `c.anchor` for `c in doc.comments`; `doc.move_comment(id, range)` | move a thread off a paragraph before deleting it |
| Tracked changes | `rdocx revision list/accept/reject --id/--author/--start-date/--end-date` | `doc.revisions`, `accept_all()`, `reject_all()`, by id, author or dates | every story: `r.story.kind` |
| Redline of two versions | `rdocx compare A B --author N --timestamp T --granularity word [--ignore-comments] -o OUT` | `a.compare(b, author, timestamp, granularity="word", ignore_comments=True)` | default: whole runs; a rebuilt TOC, a changed picture or chart shows as revisions: check |
| Table of contents | `rdocx toc rebuild F -o OUT`, `docx_ops.py toc IN OUT` | `doc.rebuild_toc()`; a new one: `doc.insert_toc(bi, max_level=3)`, then `rebuild_toc()` | |
| Fields | `rdocx fields update F -o OUT [--now 2026-10-10T09:00:00]` | `doc.update_fields()`, `update_page_fields()` | header and footer page numbers keep their cached text, as in Word |
| PDF | `rdocx convert F --to pdf [--revision-view tracked [--revision-style legal]] -o NEW.pdf` | `doc.to_pdf(revision_view="tracked", revision_style="legal")` | accepted view by default; `tracked` shows the revisions, `legal` in colour |
| PNG pages | `rdocx render F -o NEW_DIR --pages 1-3 --dpi 100 [--font-dir DIR]` | `doc.render_pages(dpi=, pages=[0, 1], font_dir=)` | CLI pages one-based, Python zero-based; `font_dir`: fonts not installed |
| Markdown, HTML, ODT, RTF, EPUB | `rdocx convert F --to md -o NEW.md` / `--to html` | `doc.text()`, `doc.to_markdown()`, `to_html()`, `to_odt()`, `to_rtf()`, `to_epub()` | no comments; ODT, RTF and EPUB return `bytes` and the `diagnostics` of what they dropped |
| New document from Markdown or HTML | `rdocx convert notes.md -o NEW.docx [--template brand.dotx]`, `rdocx convert page.html -o NEW.docx` | `rdocx.Document.from_markdown(text, template=, images=)`, `Document.from_html(html)`, `doc.insert_html(html, after=p)` | each returns what it dropped: read it |
| Charts | | `doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, w, h, after=p)`, `doc.set_chart_legend(p, visible=True, position="b")` | data as in python-pptx (`rdocx.CategoryChartData`); `doc.charts` reads them back |
| Embedded fonts | `rdocx font list/embed/remove/settings` | `doc.fonts`, `doc.embed_font(name, file, "regular", license=...)`, `remove_embedded_font(name, kind)` | rule 10 |
| Validity | `rdocx validate F` | `doc.validate()` | every related part and every style id (rule 7) |
| Sign, check signatures | `rdocx sign F --key K.pem --cert C.pem -o OUT`, `rdocx verify-signatures --json F` | `doc.sign(key, cert)`, `doc.verify_signatures()` | sign last: an edit breaks it; no chain trust |
| What changed between two files | `rdocx diff A B` | | by paragraph, in every story |
| New document | | `rdocx.Document()`, or a template .docx or .dotx emptied | Word's usual styles; a .dotx saved as .docx becomes a document |
| Combine documents, equations, a missing feature | | `doc.insert_document(other, at=bi)`, `doc.import_fragment(location, fragment)`; `p.add_equation(latex, format="latex", display=False)`; `p.xml` / `p.replace_xml(xml)` | raw XML, the way out instead of lxml: an element Word does not know raises |

Commands and full signatures: `references/cli.md`, `references/python-api.md` (and the `.pyi` stubs it
points to). Worked, tested examples for each row: `references/recipes.md`. Ready-made commands:
`scripts/docx_ops.py` (`text`, `count`, `replace`, `comment`, `toc`, `pages`). Look a name up
(`grep -n NAME references/python-api.md`) rather than reading a reference. The skill does not list every method
or flag: before concluding that rdocx cannot do something, search the build itself, `grep -n -i WORD` in its type
stubs (`_rdocx.pyi`, in the folder `$R/python -c "import rdocx, os; print(os.path.dirname(rdocx.__file__))"`
prints) and `$R/rdocx --help`, `$R/rdocx COMMAND --help`.

## When rdocx cannot do a step

When a step is blocked (`references/gaps.md` gives each known gap's workaround and fallback):

1. Search the stubs and `--help` (above), then say so in one line to the user (what is blocked, which gap).
2. Use the workaround if the gap has one; otherwise do **that step only** with the built-in `docx` skill
   (python-docx, lxml or LibreOffice), on the file rdocx saved, and keep rdocx for everything else,
   including the verification.
3. Report the gap (next section) unless `references/gaps.md` already lists it.

## Reporting a bug or a missing feature

When rdocx fails, refuses, loses data, renders wrong, or lacks an API you needed, and `references/gaps.md`
does not list it:

- Reduce it to a **neutral reproduction**: a short script that builds its own input (python-docx is fine
  for that) and prints the wrong result next to the expected one. Never include the user's
  document, its text, names or paths.
- Ask the user before writing it to their folder; then write it as `rdocx-findings/YYYY-MM-DD-<slug>.md` in
  the working folder (what, expected, observed, the script, the pinned commit from `rdocx_env.py status`),
  or give it in the reply if they prefer. Do not open upstream issues yourself.
