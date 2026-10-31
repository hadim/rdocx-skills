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

## Results (full run, 2026-10-31, pin 401e3f4218e6, Linux x86_64, 32 CPUs)

`results/speed.json`, 768 measurements in 90 minutes. Machine load: the full run
started at once (1-minute load 1.01); over the 239 groups the 1-minute load stayed between 0.94 and
1.77 (median 1.0): the other agents' LibreOffice work had ended. Median seconds of each stack's primary command,
process start included; `n/a` = the stack has no means; `error` = non-zero exit.

### Synthetic files (docx 6 / 52 / 198 pages, pptx 15 / 60 / 150 slides)

| op | rdocx S / M / L | builtin S / M / L | pylib S / M / L | builtin/rdocx S / M / L |
|---|---|---|---|---|
| docx.read | 0.051 / 0.475 / 2.22 | 0.067 / 0.419 / 1.59 | 0.052 / 0.115 / 0.348 | 1.30 / 0.881 / 0.716 |
| docx.replace | 0.630 / 5.62 / 25.7 | 0.100 / 0.418 / 1.77 | 0.059 / 0.181 / 0.614 | 0.159 / 0.074 / 0.069 |
| docx.replace+validate | 0.656 / 5.86 / 26.7 | 0.249 / 0.825 / 3.19 | n/a / n/a / n/a | 0.380 / 0.141 / 0.119 |
| docx.comment | 0.473 / 4.80 / 23.2 | 0.136 / 0.450 / 1.80 | 0.047 / 0.091 / 0.247 | 0.287 / 0.094 / 0.078 |
| docx.comment+validate | 0.500 / 5.04 / 24.1 | 0.405 / 0.969 / 3.34 | n/a / n/a / n/a | 0.810 / 0.192 / 0.138 |
| docx.accept | 0.131 / 1.18 / 4.89 | 0.405 / 0.773 / 2.24 | n/a / n/a / n/a | 3.08 / 0.655 / 0.459 |
| docx.render | 0.184 / 2.02 / 11.2 | 0.910 / 7.27 / 27.9 | n/a / n/a / n/a | 4.93 / 3.60 / 2.50 |
| docx.render-page1 | 0.091 / 0.797 / 6.35 | 0.514 / 1.16 / 3.99 | n/a / n/a / n/a | 5.64 / 1.45 / 0.628 |
| docx.pdf | 0.108 / 1.01 / 7.46 | 0.461 / 1.11 / 3.93 | n/a / n/a / n/a | 4.25 / 1.11 / 0.527 |
| docx.markdown | 0.044 / 0.374 / 1.80 | 0.067 / 0.415 / 1.58 | n/a / n/a / n/a | 1.51 / 1.11 / 0.877 |
| docx.validate | 0.028 / 0.234 / 0.971 | 0.139 / 0.370 / 1.30 | n/a / n/a / n/a | 4.94 / 1.58 / 1.34 |
| pptx.read | 0.025 / 0.114 / 0.294 | 0.081 / 0.124 / 0.214 | 0.074 / 0.111 / 0.182 | 3.32 / 1.09 / 0.728 |
| pptx.replace | 0.041 / 0.232 / 0.571 | 0.058 / 0.111 / 0.219 | 0.086 / 0.151 / 0.282 | 1.40 / 0.477 / 0.384 |
| pptx.replace+validate | 0.072 / error / error | 0.243 / 0.606 / 1.33 | n/a / n/a / n/a | 3.36 / - / - |
| pptx.duplicate | 0.033 / 0.138 / 0.347 | 0.043 / 0.093 / 0.193 | 0.083 / 0.131 / 0.230 | 1.28 / 0.675 / 0.555 |
| pptx.fit | error / error / error | n/a / n/a / n/a | n/a / n/a / n/a | - / - / - |
| pptx.render | 0.284 / 1.28 / 3.26 | 1.53 / 5.70 / 14.1 | n/a / n/a / n/a | 5.39 / 4.45 / 4.31 |
| pptx.render-slide1 | 0.147 / 0.650 / 1.64 | 0.701 / 1.63 / 3.48 | n/a / n/a / n/a | 4.76 / 2.50 / 2.13 |
| pptx.thumbnail | 0.142 / 0.645 / 1.64 | 0.929 / 2.60 / 5.94 | n/a / n/a / n/a | 6.54 / 4.03 / 3.63 |
| pptx.pdf | 0.169 / 0.807 / 2.04 | 0.641 / 1.57 / 3.42 | n/a / n/a / n/a | 3.79 / 1.94 / 1.68 |
| pptx.validate | 0.034 / 0.147 / 0.397 | 0.143 / 0.355 / 0.767 | n/a / n/a / n/a | 4.24 / 2.41 / 1.93 |
| docx.create | 4.24 / 31.9 / 208.6 | 0.114 / 0.283 / 0.723 | 0.112 / 0.470 / 1.79 | 0.027 / 0.009 / 0.003 |
| pptx.create | 0.316 / 2.26 / 7.40 | 0.145 / 0.265 / 0.481 | 0.127 / 0.287 / 0.652 | 0.460 / 0.118 / 0.065 |

The ratio is builtin over rdocx: above 1, rdocx is faster. pylib/rdocx per row is in the JSON (`summary`).

### Real files (7 Word documents rdocx opens, 5 decks): median of the per-file ratios (min to max)

| op | builtin/rdocx | pylib/rdocx |
|---|---|---|
| docx.read | 1.52 (0.18 to 2.06, 7 files) | 0.90 (0.04 to 3.78, 7 files) |
| docx.replace | 0.17 (0.06 to 0.40, 7 files) | 0.09 (0.01 to 0.28, 7 files) |
| docx.replace+validate | 0.39 (0.12 to 1.04, 7 files) | - |
| docx.comment | 0.24 (0.12 to 0.67, 6 files) | 0.08 (0.02 to 0.31, 6 files) |
| docx.comment+validate | 0.66 (0.22 to 2.16, 6 files) | - |
| docx.render | 5.13 (2.29 to 8.57, 7 files) | - |
| docx.render-page1 | 5.24 (0.56 to 10.03, 7 files) | - |
| docx.pdf | 4.16 (0.50 to 10.95, 7 files) | - |
| docx.markdown | 1.85 (0.19 to 2.33, 7 files) | - |
| docx.validate | 3.70 (1.67 to 10.28, 9 files; on the two Social Impacts reports rdocx answers "cannot open") | - |
| pptx.read | 1.46 (0.63 to 2.07, 5 files) | 1.39 (0.59 to 1.90, 5 files) |
| pptx.replace | 0.65 (0.29 to 1.01, 5 files) | 0.82 (0.46 to 1.30, 5 files) |
| pptx.replace+validate | 1.72 (1.54 to 2.18, 5 files) | - |
| pptx.duplicate | 0.88 (0.38 to 1.05, 5 files) | 1.21 (0.70 to 1.62, 5 files) |
| pptx.render | 3.05 (1.86 to 6.05, 5 files) | - |
| pptx.render-slide1 | 1.67 (1.04 to 5.16, 5 files) | - |
| pptx.thumbnail | 2.48 (1.53 to 6.66, 5 files) | - |
| pptx.pdf | 1.34 (0.84 to 4.36, 5 files) | - |
| pptx.validate | 2.76 (2.21 to 2.94, 5 files) | - |

### Render time per page or slide

| op | rdocx s/page S / M / L | builtin s/page S / M / L |
|---|---|---|
| docx.render | 0.031 / 0.040 / 0.057 | 0.152 / 0.140 / 0.141 |
| pptx.render | 0.019 / 0.021 / 0.022 | 0.102 / 0.095 / 0.094 |

(The built-in render is LibreOffice to PDF, then `pdftoppm`: on docx-L, LibreOffice's PDF takes 3.9 s and `pdftoppm`
24 s of the 27.9 s.)

### Variants

- Word, replace (docx-M / docx-L): `docx_ops.py replace` 5.62 / 25.7 s, `rdocx replace --expect` 0.70 / 2.76 s,
  built-in 0.42 / 1.77 s, python-docx 0.18 / 0.61 s. The skill's helper costs 8 to 9 times the CLI (finding 7).
- Word, comment (docx-M / docx-L): `docx_ops.py comment` 4.80 / 23.2 s, `rdocx comment add --anchor` 0.76 / 3.07 s,
  built-in 0.45 / 1.80 s, python-docx 0.09 / 0.25 s. `comment_on_text` runs the whole edit twice (on a probe copy,
  read back from bytes, then for real).
- Word, read (docx-M / docx-L): `rdocx text --json` 0.48 / 2.22 s, pandoc 0.42 / 1.59 s, the built-in's own lxml
  read 0.07 / 0.20 s, python-docx 0.11 / 0.35 s. On the Taylor Review interim evaluation (66 pages, 11 charts)
  `rdocx text --json` takes 3.9 s against 0.7 s for pandoc.
- Word, create docx-S, naive rdocx (one body, cell paragraphs styled through their handles): 26.4 s against 4.2 s
  for the chaptered script, 0.11 s for docx-js and python-docx.
- PowerPoint, fit: `rpptx fit --json` fails on 8 of 9 decks (finding 5); on the python-pptx deck 0.93 s (1,453
  frames). `pptx_ops.py overflow` works everywhere: 0.17 / 0.69 / 1.74 s on pptx-S / M / L, 1.0 to 1.7 s on the real
  decks.
- PowerPoint, thumbnail: `rpptx thumbnail` (slide 1, 320 px) 0.14 / 0.65 / 1.64 s against `thumbnail.py` (a grid of
  every slide, 1 to 13 images) 0.93 / 2.60 / 5.94 s: not the same output.

### Where rdocx is slower

- Every Word edit through the skill's helpers: replace 6 to 15 times the built-in workflow and 11 to 42 times
  python-docx on the synthetic files, 2.5 to 17 times the built-in on the real ones; comment 3.5 to 13 times the
  built-in. With the validation step on both sides the built-in path is still 5 to 8.4 times faster on M and L.
  The rdocx CLI (`rdocx replace`, `rdocx comment add`) is 1.6 to 1.7 times the built-in workflow.
- Creating a Word document: 4.2 / 31.9 / 209 s against 0.11 / 0.28 / 0.72 s (docx-js) and 0.11 / 0.47 / 1.79 s
  (python-docx), 37 to 290 times slower; most of it in table cells (finding 1). Decks: 0.32 / 2.26 / 7.40 s against
  0.15 / 0.27 / 0.48 s (pptxgenjs), 2 to 15 times slower.
- Large Word files with LibreOffice as the other side: PDF of docx-L 7.5 s against 3.9 s (0.53); page 1 of docx-L 6.4 s
  against 4.0 s, since rdocx lays out the whole document first; accept all revisions 1.2 / 4.9 s against 0.8 / 2.2 s
  (docx-M / L); the two Taylor Review reports: PDF 0.5 to 0.7, Markdown 0.19 to 0.36, read 0.18 to 0.33.
- Reading large files: `rdocx text --json` on docx-L 2.2 s against pandoc 1.6 s and python-docx 0.35 s;
  `rpptx text --notes` on pptx-L 0.29 s against 0.21 / 0.18 s.
- Deck edits on M and L: replace 0.23 / 0.57 s against 0.11 / 0.22 s (built-in), duplicate 0.14 / 0.35 s against
  0.09 / 0.19 s.

### Where rdocx is faster

- Rendering: Word pages 2.5 to 5 times faster than LibreOffice + pdftoppm on the synthetic files, 2.3 to 8.6 times on
  the real ones (0.03 to 0.06 s a page against 0.14 to 0.15); slides 4.3 to 5.4 times (0.02 s a slide against 0.1).
- Deck PDF 1.7 to 3.8 times, small and medium Word PDF 1.1 to 4.3 times, validation 1.3 to 5 times, every
  small-file command (process start: a 6-page document or 15-slide deck answers in 0.03 to 0.2 s).

### Correctness and validity

- Every primary rdocx command whose output was checked passed its check; the 32 rdocx errors are the two Social Impacts
  reports (22, duplicate bookmark ids), `rpptx fit` (8, finding 5), and `rpptx replace` + `rpptx validate` on pptx-M
  and pptx-L (2: the pptxgenjs deck's own duplicate shape id fails validation before and after; `rpptx validate` has
  no `--original` baseline). Every built-in and python-docx command passed its check.
- Validity of outputs: rdocx's created reports fail the built-in `validate.py` on `word/footnotes.xml` (finding 11);
  `rpptx slide duplicate` on the pptxgenjs decks fails it on the theme sharing (finding 10); the built-in comment
  (`comment.py`) and python-docx comment outputs fail `rdocx validate` (the comments part uses a `CommentReference`
  style the document does not define); python-docx's created reports fail both (undefined `CommentReference`, its
  template's `w:zoom` without `w:percent`); pptxgenjs's M and L decks fail `rpptx validate` (duplicate shape id,
  corpus/NOTES.md), and so does every edit of them.
- Created files differ from the corpus file of the same plan where a library could not follow: rpptx 731 paragraphs
  against 641 (chart titles moved into the commentary, the default template's layouts: 11 against 8), python-pptx no
  sections (0 against 6), python-docx 7,775 `w:p` against 8,077 (no TOC entries until Word updates the field).
- pandoc drops the Title and Subtitle paragraphs into metadata (`-t markdown` without `--standalone`); `add_slide.py`
  does not copy speaker notes (`notes_copied: false`).

## rdocx findings

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
