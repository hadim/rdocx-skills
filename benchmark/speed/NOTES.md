# Speed: notes

How long do the operations an agent actually runs take, on realistic files, per stack (`rdocx`, `builtin`, `pylib`,
README.md)?

```bash
.cache/venv/bin/python speed/run.py                     # full run, results/speed.json (about 1.5 h on 32 CPUs)
.cache/venv/bin/python speed/run.py --quick             # 1 run per command, no warm-up: .cache/work/speed/quick.json
.cache/venv/bin/python speed/run.py --quick --only 'docx.replace:docx-M:*,pptx.*:pptx-S:*'   # op:file:stack:variant
```

The full run waits for a quiet machine first (1-minute load under 4, polled every 2 minutes, up to 30 minutes;
`--no-wait` skips it) and records the load before and after every group of commands (one operation on one file,
all stacks) in `load` and in each measurement.

## Files

| File | What |
|---|---|
| `run.py` | runner: inputs, commands, timing, checks, validity, summary |
| `oracle.py` | neutral reads (zipfile + lxml): word, phrase and sample selection, output checks |
| `docx_builtin_*.sh/.py`, `pptx_builtin_replace.*`, `lo_render.sh` | the built-in stack's workflows (unzip, skill scripts, lxml edit, zip; LibreOffice + pdftoppm) |
| `docx_pylib_*.py`, `pptx_pylib_*.py` | python-docx / python-pptx scripts |
| `docx_create_rdocx.py`, `docx_create_pylib.py`, `pptx_create_rdocx.py` | creation from the corpus plans (the built-in stack and python-pptx use the corpus's own generators) |
| `rc0.sh` | runs a checker whose exit code 1 is an answer (validate, fit), not an error |

## Method

- **Inputs.** Every corpus file (`corpus/manifest.json`): the synthetic docx-S/M/L (6/52/198 LibreOffice pages),
  docx-M-clean, docx-M-lo, pptx-S/M/L (15/60/150 slides), pptx-M-pythonpptx, the 9 real GOV.UK / data.gouv.fr
  documents (2 to 153 pages) and 5 real decks (13 to 69 slides). Creation uses the plans in `.cache/corpus/plans/`.
- **Commands.** Each measurement is the command an agent runs, process start included (`bash` scripts for the
  multi-step workflows, so each step's process start counts). Environments come from `common.env_*()`; `TMPDIR`
  points into `.cache/work/speed/tmp`, emptied before every run (untimed), as is every output.
- **Timing.** `common.timeit`: a warm-up run (timed, kept as `warmup`), then the median of 5 runs; 3 runs for
  LibreOffice operations and for anything whose warm-up took over 20 s; 1 run when it took over 60 s (flagged
  `single-run`). `--quick`: 1 run, no warm-up.
- **Correctness.** Each output is checked cheaply by `oracle.py`, never with a stack's reader: reads must return
  three sample phrases spread over the document (and a notes sample for decks); a replacement must leave no
  occurrence of the word in the story parts and add exactly N of the new text; a comment must exist with its text,
  anchored exactly on the phrase (the `w:t` between `commentRangeStart` and `commentRangeEnd`) with its
  reference; accepting must leave no revision mark in any story part; renders must give one PNG per page or slide
  of that renderer (rdocx: the pages of its own PDF of the file, `rdocx convert --font-dir`, finding 8;
  LibreOffice: `pages.libreoffice` of the manifest; decks: the slide count); a PDF must have pages (`pdfinfo`;
  `pages_match_manifest` records whether they match the manifest's count); a duplicated slide must repeat the source's text right after it; created files are counted
  with `corpus/ooxml_stats.py` and compared with the corpus file built from the same plan. `status` is `ok`, `fail`
  (ran but the check failed: a fast wrong answer is not a win), `error` (non-zero exit) or `unsupported`.
- **Validity.** Every Office output is validated by `rdocx validate` / `rpptx validate` and the built-in
  `validate.py` (`--original IN` for an edit), untimed; recorded in `valid`.
- **Summary.** `summary` has one row per operation and file: the median of each stack's primary command, its
  status, `builtin/rdocx` and `pylib/rdocx`, and every variant's median.

## Operations

| op | rdocx | builtin | pylib |
|---|---|---|---|
| docx.read | `rdocx text --json` | `pandoc -t markdown` (variant: unzip + lxml read of every story, as features D02) | python-docx: body paragraphs, tables, headers, footers |
| docx.replace | `docx_ops.py replace IN OUT --edit OLD NEW N` (variant: `rdocx replace --expect N`) | unzip, `merge_runs.py`, lxml edit of document/headers/footers, zip | python-docx run-level replace |
| docx.replace+validate | the same, then `rdocx validate OUT` | the same, then `validate.py OUT --original IN` | n/a (no validator) |
| docx.comment(+validate) | `docx_ops.py comment --anchor PHRASE --date ...` (variant: `rdocx comment add --anchor`) | unzip, `merge_runs.py`, `comment.py` (directory mode), lxml marker insertion, zip | python-docx 1.2 `add_comment` on split runs |
| docx.accept | `rdocx revision accept` | `accept_changes.py` (LibreOffice) | n/a |
| docx.render / render-page1 | `rdocx render --dpi 100 --font-dir` (`--pages 1`) | `soffice.py` to PDF, then `pdftoppm -png -r 100` (`-f 1 -l 1`) | n/a |
| docx.pdf | `rdocx convert --to pdf --font-dir` | `soffice.py --convert-to pdf` | n/a |
| docx.markdown | `rdocx convert --to md` | `pandoc -t markdown -o` | n/a |
| docx.validate | `rdocx validate` | `validate.py` | n/a |
| docx.create | `docx_create_rdocx.py` (rdocx Python) | `corpus/gen_docx.js` (docx-js, the corpus generator) | `docx_create_pylib.py` |
| pptx.read | `rpptx text --notes` | `slide_text.py` | python-pptx |
| pptx.replace(+validate) | `rpptx replace --expect N` (then `rpptx validate`) | zipfile unpack, lxml edit, zip (then `validate.py --original`) | python-pptx run-level replace |
| pptx.duplicate | `rpptx slide duplicate` (the middle slide) | `add_slide.py IN slideK.xml --after slideK.xml -o OUT` | python-pptx + lxml copy (no API) |
| pptx.fit | `rpptx fit --json` (variant: `pptx_ops.py overflow`) | n/a (visual QA) | n/a |
| pptx.render / render-slide1 | `rpptx render --dpi 100 --font-dir` (`--slide 1`) | `soffice.py` + `pdftoppm` | n/a |
| pptx.thumbnail | `rpptx thumbnail` (slide 1 at 320 px: not a grid) | `thumbnail.py` (labelled grid of every slide) | n/a |
| pptx.pdf | `rpptx convert --to pdf --font-dir` | `soffice.py --convert-to pdf` | n/a |
| pptx.validate | `rpptx validate` | `validate.py` | n/a |
| pptx.create | `pptx_create_rdocx.py` (rpptx Python) | `corpus/gen_pptx.js` (pptxgenjs + `apply_theme.js`) | `corpus/gen_pptx_pythonpptx.py` |

## Judgement calls

- **Which word, which phrase.** `oracle.choose_docx_word` / `choose_pptx_word` take the most frequent word of five
  letters or more that occurs only where every stack's documented replacement reaches it (plain runs of body
  paragraphs and table cells; for decks, text runs of slides and of the notes body), and nowhere else in the
  package (headers, notes, comments, charts, text boxes, links, fields, tracked changes), so that a time is
  compared with a time; the replacement is the word in capitals and N is that count. The phrase for a comment is
  four ASCII words occurring once, in a plain body paragraph near the middle of the document, across a run boundary
  when one is available. The charity examiner template has no such paragraph (its text is in tables): no comment
  measurement there. The hard cases (split runs, tables, headers, text boxes) are the features section's job.
- **Validation step timed apart.** The built-in skill's editing workflow ends with `validate.py --original`;
  the rdocx skill's with `rdocx validate` (rule 7). Both are measured with and without that step
  (`*+validate`); the primary comparison is without.
- **Renders.** The built-in path renders by converting the whole document to PDF and then rasterising with
  `pdftoppm` (PNG, to match rdocx; the skill's docx example uses `-jpeg`). "Page 1 only" still converts the whole
  document with LibreOffice: that is the skill's path. LibreOffice shows tracked changes in the render (docx-S/M/L,
  docx-M-lo carry revisions), rdocx the accepted view: same pages, different marks. The PNG count is checked
  against each renderer's own page count (rdocx paginates docx-L to 195 pages, LibreOffice to 198).
- **Creation, rdocx.** `docx_create_rdocx.py` follows the rdocx skill's advice for large documents
  (references/gaps.md: an edit costs more in a larger body; build parts aside, insert them with one
  `insert_document`): one small document per chapter, inserted in order. Tracked insertions and deletions are
  written as paragraph XML through `replace_xml` (rdocx has no API that writes a revision); the table cells'
  paragraph styles through one `replace_xml` per table (finding 1). The naive variant (everything appended to one
  body, cell paragraphs styled through their handles) is measured on docx-S only: 27 s for 6 pages.
- **Creation, built-in and pylib decks.** The corpus generators are the built-in stack's (docx-js, pptxgenjs +
  `apply_theme.js`) and python-pptx's creation paths, run from `corpus/` unchanged (they load `deterministic.js`,
  which fixes the clock and the random ids; negligible cost). `docx_create_pylib.py` is new: python-docx's API
  where it has one, its `OxmlElement` / `parse_xml` where it has none.
- **What each creator could not include** (recorded per file in `check.differs_from_corpus`): rpptx writes no
  chart title and no data labels (no API; the chart title goes into the slide's commentary text instead), and
  uses its default template's layouts (the pptxgenjs deck defines seven of its own); python-pptx has no section
  API (no sections) and its default template is 4:3; python-docx and rdocx write the TOC field and rdocx fills it
  (`rebuild_toc`) while docx-js and python-docx leave it for Word to update.
- **Fit.** `rpptx fit` exits 1 when a frame overflows (an answer): `rc0.sh` records the exit code and fails only
  on 2 or more. The built-in skill has no fit check (its way is visual QA of the renders) and python-pptx no
  layout engine: both `unsupported`.
- **Thumbnail.** `rpptx thumbnail` renders slide 1 at 320 px; the built-in `thumbnail.py` a labelled grid of every
  slide. Not the same output: recorded with a note, not a like-for-like ratio.
- **Duplicate.** `add_slide.py` does not copy the speaker notes (recorded in `check.notes_copied`); the python-pptx
  recipe copies shapes, relationships and the notes text, then moves the copy after its source.
- **python-docx merged cells.** `row.cells` repeats a merged cell in every grid column it spans: the pylib scripts
  visit each `w:tc` once (keyed on the element, not on `id()`, which lxml proxies reuse).

## Caveats

- Noisy machine during development (three other agents ran LibreOffice-heavy work); the full run itself was quiet
  (1-minute load 0.94 to 1.77). The load before and after each group is in the JSON; a re-run on a busier machine
  will show it there. LibreOffice operations have 3 runs, not 5; docx-L creation by rdocx has 1 (warm-up over 60 s).
- Single machine, Linux; LibreOffice starts a fresh profile on every call (soffice.py), as an agent pays it.
- pandoc moves Title and Subtitle paragraphs into its metadata, which `-t markdown` without `--standalone` does
  not print: the read samples are taken from other paragraphs (the title is not counted against pandoc).
- Two real reports (Social Impacts, main report and appendices) do not open in rdocx (duplicate bookmark ids,
  corpus/NOTES.md finding 1): every rdocx operation on them is an `error`, and they are left out of the ratios.

## Results

The tables of `benchmark/README.md`, generated from `results/speed.json` by `report.py`. This file does not repeat the numbers: they change with each run.

## rdocx findings

Met on the first run (2026-10-31, pin `401e3f42`). They were reported upstream as tensorbee/rdocx#431 to #437, and the pin the benchmark now runs on carries the fixes of all but the open gaps of `tests/gaps.py`. Kept as the record of what the benchmark found; the README has the current results.

Pinned build 401e3f42 (rdocx 0.16.4, rpptx 0.14.4). Each reproduction builds its own input; run with `$R/python`
(`R=~/.local/share/rdocx-skills/current/bin`), node scripts with `NODE_PATH=.cache/node/node_modules`.

1. **Every edit through a table cell's paragraph or run handle costs time that grows with the square of the
   table's cells.** `cell.paragraphs[0].style = ...`, `.alignment`, `.paragraph_format.space_after`, fetching
   `p.runs[0]` and `run.font.size` each take about 5 ms in a 30-cell table, 20 to 35 ms at 60 cells and 90 to 170
   ms at 120 cells (a 20 x 6 table formatted cell by cell: about 17 s per property); `cell.text` and
   `cell.vertical_alignment` grow linearly per call (4, 7, 13 ms), so a table filled cell by cell costs its
   cells squared. The skill says walking table handles is linear. Effect here: the corpus report built straight
   with the API takes 27 s for docx-S (6 pages); with the cell styles written by one `replace_xml` per table and
   chapters built aside, docx-L still spends 170 of its 214 s in `cell.text`, `cell.vertical_alignment` and
   `cell.shading` (cProfile). Expected: constant or linear per call.

   ```python
   import time, rdocx
   doc = rdocx.Document()
   doc.add_style("TT", based_on="Normal", font_size=rdocx.Pt(9))
   for rows in (5, 10, 20):
       t = doc.add_table(rows, 6)
       cells = [t.cell(r, c) for r in range(rows) for c in range(6)]
       for c in cells:
           c.text = "x"
       paras = [c.paragraphs[0] for c in cells]
       t0 = time.perf_counter()
       for p in paras:
           p.style = "TT"
       print(f"{len(cells)} cells: {1000 * (time.perf_counter() - t0) / len(cells):.1f} ms per style set")
   # 30 cells: 4.4 ms, 60 cells: 19.4 ms, 120 cells: 93.5 ms (another run: 5.3, 27.3, 137.0); expected: flat
   ```

2. **`rebuild_toc()` refuses a TOC that `insert_toc()` wrote before the headings existed**: `RdocxError: table of
   contents result must span paragraph boundaries`. `insert_toc` writes the field (begin, separate, end) inside one
   paragraph when there is no heading yet; `rebuild_toc` cannot fill it later. Inserting the TOC after the headings
   works. The natural order when building a report from the top (title, contents, chapters) hits it.

   ```python
   import rdocx
   doc = rdocx.Document()
   doc.add_paragraph("Contents")
   doc.insert_toc(1, max_level=3)
   doc.add_heading("One", 1)
   doc.rebuild_toc()   # RdocxError: table of contents result must span paragraph boundaries (expected: 1 entry)
   ```

3. **A numbering definition id goes stale.** Once paragraphs use instances of two definitions, rdocx renumbers the
   abstract definitions by first use; the id `add_numbering_definition` returned keeps working until an edit that
   retires every handle (a footnote, a comment, a picture, a section break), after which
   `add_numbering_instance(old_id)` makes an instance of the *other* definition: a bullet list comes out numbered
   (or the reverse, and a later call fails with `document content references numbering instance 3 at missing
   level 1`). Workaround used: every instance made up front.

   ```python
   import re, rdocx
   doc = rdocx.Document()
   bullets = doc.add_numbering_definition([rdocx.ListLevel(format="bullet", text="o"), rdocx.ListLevel(format="bullet", text="-")])
   numbers = doc.add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.")])
   n1 = doc.add_numbering_instance(numbers); doc.add_paragraph("n", style="List Paragraph").numbering = (n1, 0)
   b1 = doc.add_numbering_instance(bullets); doc.add_paragraph("b", style="List Paragraph").numbering = (b1, 1)
   doc.add_footnote(doc.add_paragraph("text"), "a note")
   b2 = doc.add_numbering_instance(bullets)
   xml = doc.part_xml("word/numbering.xml")
   fmt = dict(re.findall(r'<w:abstractNum w:abstractNumId="(\d+)".*?<w:numFmt w:val="(\w+)"', xml))
   num = dict(re.findall(r'<w:num w:numId="(\d+)"><w:abstractNumId w:val="(\d+)"/>', xml))
   print(fmt[num[str(b2)]])   # decimal (expected: bullet); without the footnote line: bullet
   ```

4. **`insert_document` merges equal list instances**, so numbered lists that restart in the inserted document
   continue each other in the target (1, 2 / 3, 4 / 5, 6 instead of 1, 2 three times), even when the target already
   has those exact instances. Workaround used: set each numbered paragraph's `numbering` again after the insert.

   ```python
   import rdocx
   base = rdocx.Document()
   numbers = base.add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.")])
   ids = [base.add_numbering_instance(numbers, start=1) for _ in range(3)]
   data = base.to_bytes()
   doc = rdocx.Document.from_bytes(data)
   for k, nid in enumerate(ids):
       part = rdocx.Document.from_bytes(data)
       for j in range(2):
           part.add_paragraph(f"list {k} item {j}", style="List Paragraph").numbering = (nid, 0)
       doc.insert_document(part, at=len(doc.paragraphs))
   print({p.numbering for p in doc.paragraphs})   # {(1, 0)} (expected: three instances, as in the parts)
   ```

5. **`rpptx fit` refuses 8 of the 9 corpus decks (exit 2)**, all five PowerPoint decks and the three pptxgenjs ones:
   `fit cannot assess mandatory owners: /ppt/notesSlides/notesSlide1.xml [2]:` (with an empty reason; any pptxgenjs
   deck, whose notes slides it always writes, and PowerPoint decks), the same for layouts and charts
   (`/ppt/slideLayouts/slideLayout13.xml [3]:`, `/ppt/charts/chart10.xml [2]:`), `presentation render failed:
   autofit /ppt/slides/slide24.xml: duplicate shape ID 25` (pptxgenjs's duplicate id, which `rpptx validate`
   reports) and `text layout failed` on a layout's table cell. Only the python-pptx deck passes. Python's
   `prs.fit_report()` and `text_layout()` (`pptx_ops.py overflow`) work on the same decks, so the agent's first
   command for rule 5 fails where the API would answer. Expected: frames it cannot assess reported as such, the
   others assessed, exit 0 or 1.

   ```js
   // NODE_PATH=.cache/node/node_modules node fit_repro.js fit.pptx ; $R/rpptx fit fit.pptx ; echo $?   -> 2
   const pptxgen = require("pptxgenjs");
   const pres = new pptxgen();
   pres.addSlide().addText("Short text", { x: 1, y: 1, w: 4, h: 1 });
   pres.writeFile({ fileName: process.argv[2] });
   // Error: fit cannot assess mandatory owners: /ppt/notesSlides/notesSlide1.xml [2]:   (exit 2)
   // $R/python -c "import rpptx; print(len(rpptx.Presentation('fit.pptx').fit_report().frames))"  -> 9 frames
   ```

6. **`rdocx replace` (and `try_replace_text`) also replaces the cached labels of Word charts**, which `rdocx text`,
   `docx_ops.py count` and the skill's list of what a replacement reaches do not show: `--expect N` taken from
   the text is refused (`expected 112 replacement(s) of "engagement", found 120` on the Taylor Review interim
   evaluation, whose 4 charts hold 8 labels), and the chart's cache no longer matches its workbook. Expected:
   charts left alone (or counted in `rdocx text` and documented).

   ```python
   import os, subprocess, rdocx
   R = os.path.expanduser("~/.local/share/rdocx-skills/current/bin/rdocx")
   doc = rdocx.Document()
   doc.add_paragraph("Sales in the North region rose.")
   data = rdocx.CategoryChartData(); data.categories = ["North", "South"]; data.add_series("Sales", [3, 4])
   doc.add_chart(rdocx.XL_CHART_TYPE.COLUMN_CLUSTERED, data, rdocx.Inches(4), rdocx.Inches(3))
   doc.save("chart.docx")
   print(subprocess.run([R, "text", "chart.docx"], capture_output=True, text=True).stdout.count("North"))   # 1
   r = subprocess.run([R, "replace", "chart.docx", "-p", "North", "-v", "NORTH", "--expect", "1", "-o", "out.docx"],
                      capture_output=True, text=True)
   print(r.returncode, r.stderr)   # 1 Error: pair 0: expected 1 replacement(s) of "North", found 2
   ```

7. **Python's `try_replace_text` is slower than the CLI's replacement, and `docx_ops.py replace` runs it twice.** On
   the Taylor Review past-schemes report (47 pages, 88 occurrences): `Document.try_replace_text` alone 4.75 s, the
   whole `rdocx replace --expect 88` command 1.48 s; `docx_ops.replace_batch` (open 0.9 s, a probe copy 1.2 s,
   `try_replace_text` on the probe and on the document, 9.5 s, two XML scans 0.9 s, save 0.3 s) 13 s. On a plain
   2,000-paragraph python-docx file the gap is small (0.45 s against 0.32 s for the whole CLI command), so it
   depends on the document (tables, footnotes, fields); not reduced further.

8. **The CLI and the Python binding paginate the same file differently with the same font folder.** Taylor Review
   past-schemes report (corpus/real.json, pinned by SHA-256; Word and its PDF twin: 47 pages):
   `doc.page_count(font_dir=".cache/fonts")`, `doc.layout(font_dir=...)` and `len(doc.render_pages(dpi=100,
   font_dir=...))` give 47, while `rdocx render F -o DIR --dpi 100 --font-dir .cache/fonts` writes 48 PNGs and
   `rdocx convert F --to pdf --font-dir .cache/fonts` a 48-page PDF, as Python does without `font_dir` (48). The
   `--font-dir` flag seems not to reach the CLI's layout the way `font_dir=` reaches Python's. Expected: one
   pagination for one document and one font folder. Reproduction (no rdocx call changes the file):

   ```bash
   F=.cache/corpus/real/gov-uk-dcms-taylor-review-past-schemes/office/Report_on_Past_Scheme_Evaluations__Word_.docx
   $R/rdocx render "$F" -o /tmp/r48 --dpi 20 --font-dir .cache/fonts && ls /tmp/r48 | wc -l        # 48
   $R/python -c "import rdocx, sys; d = rdocx.Document.open(sys.argv[1]); print(d.page_count(font_dir='.cache/fonts'), len(d.render_pages(dpi=20, font_dir='.cache/fonts')))" "$F"   # 47 47
   ```

   The speed checks therefore compare rdocx's PNG count with the page count of rdocx's own PDF (CLI), not with the
   manifest's `pages.rdocx` (Python).


9. Known: two real reports do not open (duplicate bookmark ids), corpus/NOTES.md finding 1; every rdocx command on
   them fails, `docx.read` included.

10. **`rpptx slide duplicate` reorders `<p:presentation>`'s children on a pptxgenjs deck** (`notesMasterIdLst` moved
    before `sldIdLst`, the schema's order). pptxgenjs writes it after `sldIdLst` and points the notes master and the
    slide master at one theme part; the built-in pptx skill states that PowerPoint reads that as written and refuses
    the deck once the element is moved, and its `validate.py` fails the duplicated deck (`Found 1 master(s) sharing a
    theme part ... move <p:notesMasterIdLst> back to directly after <p:sldIdLst>`). Not checked in PowerPoint here
    (no PowerPoint on this machine): the claim is the built-in skill's. `rpptx replace` keeps the order; `add_slide.py`
    and the python-pptx copy keep it too. Reproduction: any pptxgenjs deck of two slides (the fit repro above with a
    second `addSlide()`), `$R/rpptx slide duplicate deck.pptx 1 -o dup.pptx`, then
    `unzip -p dup.pptx ppt/presentation.xml` shows `<p:notesMasterIdLst>` before `<p:sldIdLst>`; the input has it after.

11. Known from the features section (D11): `add_footnote` writes the note's reference run as
    `<w:r><w:footnoteRef/><w:rPr>...</w:rPr></w:r>` (properties after the content, against the schema) and no
    separator notes; `validate.py` fails every rdocx-created report here on `word/footnotes.xml`, `rdocx validate`
    passes it.
