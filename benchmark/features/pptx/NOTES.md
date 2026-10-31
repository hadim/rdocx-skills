# features/pptx: method, judgement calls, caveats, rdocx findings

Fifteen PowerPoint tasks, each done three ways (`rdocx` = rpptx 0.14.4 at pin 401e3f42, `builtin` = the built-in
pptx skill's documented workflow, `pylib` = python-pptx 1.0.2 with lxml where it has no API), checked by one
neutral oracle. `run.py` regenerates `results/features_pptx.json` from scratch (about one minute; it can be run by
task groups with `--tasks P01,P02 --merge`).

## Layout

- `lib/inputs.py`: the deterministic input decks (used by `make_input.py` only). python-pptx's default master and
  eleven layouts, widened to 16:9, a theme named Northwind (12 colours, Cambria headings, Calibri body: rendered
  with Caladea and Carlito), a logo shape on the master, placeholders, runs with mixed bold/italic/colour/underline,
  notes, groups, tables (Medium Style 2), native charts with embedded workbooks, generated PNG pictures. A second
  theme, Harbor, with renamed layouts serves P05. Zips are rewritten with fixed timestamps: same bytes every run.
- `lib/oracle.py`: the oracle's shared reader (`zipfile` and `lxml` only): parts, relationships, content types,
  `sldIdLst`, layouts, masters, themes, notes, placeholders, text with fields and breaks; `integrity()` (every XML
  part parses, every internal relationship target exists, every part has a content type, no repeated relationship
  id; an override naming a missing part is only a warning, since pptxgenjs writes them and Office accepts them);
  `orphans()` (parts no relationship chain from the package reaches); the validators.
- `Pxx-slug/`: `task.md` (one line), `make_input.py OUTDIR`, `rdocx.py|rdocx.sh`, `builtin.py|builtin.sh|builtin.js`,
  `pylib.py` (or `pylib.unsupported` with the reason), `check.py OUTPUT --inputs DIR [--original F] [--no-validate]`.
  Implementations take their input and output paths as arguments; the built-in skill's folder comes from
  `PPTX_SKILL`, the benchmark fonts from `FONT_DIR` (run.py sets both).

## Method

- **Status.** `pass` when the implementation exits 0 and `check.py` finds nothing wrong; `fail` otherwise (the
  oracle's errors are in `notes`); `unsupported` when the stack has no means (one case: python-pptx cannot render).
  The oracle checks what the task asks, plus package integrity on every .pptx output.
- **Validators** (`valid`, .pptx outputs only) are reported, not folded into the status: `rpptx_validate` (`rpptx
  validate`), `builtin_validate` (the skill's `scripts/office/validate.py`, with `--original INPUT` for every edit
  and for the template-derived P15, as the skill says), `libreoffice_opens` (headless conversion to PDF with its
  own profile, at least one page). PowerPoint itself is not available here.
- **LOC**: non-blank, non-comment lines of the entry file (and of any extra file its `# bench:` header names; none
  needed one). The built-in skill's own scripts and rpptx's `pptx_ops.py` are not counted (they are the stack).
- **processes**: the programs the implementation starts: the interpreter (or bash) plus every python, node,
  soffice, pdftoppm or rpptx it launches; coreutils (`mktemp`, `mkdir`, `rm`) and a tool's own children
  (soffice.bin) not counted. Declared in each file's `# bench:` header.
- **raw_xml**: the implementation writes or edits OOXML itself (lxml, string XML, private `_element` / `_sldIdLst`
  access). Reading XML for a value does not count.
- **seconds**: wall clock of the implementation's command as an agent runs it, process start included, median of
  5 runs after one warm-up (3 when it runs LibreOffice: P06 and P12 builtin), each run into a fresh output folder.
  The QA an agent would add afterwards (validate, fit, render) is excluded for every stack; it is what `valid`
  measures.
- **No strawman.** Each implementation takes the shortest path its skill documents: rdocx the CLI for one-shot
  whole-deck operations (`replace`, `footer`, `slide import`, `render`, `convert`) and the Python binding with
  `pptx_ops.save_atomic` for multi-step edits (as its recipes do); builtin pptxgenjs as a structured deck with
  `apply_theme.js` for creation, unzip / lxml / zip with `add_slide.py`, `clean.py`, `slide_text.py` and
  `soffice.py` + `pdftoppm` otherwise; pylib python-pptx's public API first.

## Judgement calls

- **P01**: closing slide on the Section Header layout in every stack, so that "numbers on all slides but the title
  slide" and the Header and Footer dialog's "don't show on title slide" agree. pptxgenjs: 6 layouts via
  `defineSlideMaster` with named placeholders and `slideNumber`, scheme colours, sections, `applyTheme` after
  `writeFile`. python-pptx's default template is 4:3; pylib sets 16:9 but leaves the layouts' 10-inch geometry (a
  visual defect the oracle does not score: content sits in the left three quarters).
- **P02**: the builtin JSON comes from parsing `slide_text.py`'s Markdown, the skill's documented reader.
- **P03**: expected count is an argument; the refusal half reruns the same command with 8 and requires a non-zero
  exit and no output file. The oracle compares each paragraph character by character: text outside a match keeps
  its run properties; the replacement may take any matched run's properties (rpptx and the hand-written code both
  give it the first run's).
- **P04**: slides are named by their original numbers (duplicate 3 after 5, move 8 to position 2, delete 4, hide
  6), the reading where "slide 4" is Delta, the slide with its own picture and notes, so that deletion has orphans
  to clean. "Duplicate" includes the notes, as in PowerPoint: `add_slide.py` drops them (documented), so builtin
  copies the notes slide by hand.
- **P05**: notes are not required. rpptx brings them; builtin and pylib leave them behind (bringing them needs a
  notes master, which deck A lacks). `oracle_info.b2_notes_kept` records it.
- **P06**: truth is "visibly overflowing in PowerPoint", with every overflow gross and every fit at 13 % slack or
  more; autofit is off in the master and on the text boxes, so no renderer may shrink anything. One fitting frame
  is a table whose cell is taller than its stored row height: PowerPoint, LibreOffice and rpptx's renderer all
  grow the row and the table stays on the slide, so it does not overflow. All three stacks report it (precision
  5/6, recall 1); with "cell text taller than the stored row" as the truth instead, all three would score 1/1. The
  builtin way is visual QA; the script stand-in renders with `soffice.py` + `pdftoppm` (50 dpi) and counts ink in
  a 0.4-inch band around each text frame that no other shape covers. It is blind to text running into a
  neighbouring shape and reads any row growth as overflow; the input's spacing is ordinary, not tuned to it.
  pylib estimates heights itself (Pillow font metrics on Carlito/Caladea, sizes inherited through layout and
  master with lxml).
- **P07**: modern comments (PowerPoint 365 threads, `p188`) are the only format with a shape anchor and replies, so
  builtin and pylib write them by hand (authors part, comment part, the slide's `p188:commentRel` extension, rels,
  content types). A legacy comment would be reported as "not anchored".
- **P08**: the builtin rewrites caches and formulas with lxml and the embedded workbook with openpyxl (installed in
  the stack's environment; the pptx skill names no tool for workbooks).
- **P09**: the old image is shared with slide 2 (one media part), so overwriting `ppt/media/image1.png` would be
  wrong; builtin adds a new media part.
- **P12**: besides count, size and non-blank, the oracle checks that the two non-text elements are drawn: bars in
  the series colour inside slide 6's chart frame, the photo's saturated pixels on slide 7 (standard-library PNG
  decoding). Placement fidelity is the visual section's subject, not checked here.
- **P13**: list nesting is part of "bullets". `slide_text.py` prints paragraphs without their level, so the
  builtin's Markdown cannot nest; an agent that writes python-pptx code instead gets the pylib result. rpptx's
  Markdown is decoded as CommonMark (entities, backslash escapes, inline HTML) before matching.
- **P15**: builtin picks the layouts by name in the script; `thumbnail.py` (choosing layouts by eye) has no role in
  a script. The picture placeholder is filled with a computed `srcRect` (cropped to its 16:9 frame), as PowerPoint
  and python-pptx do; the oracle rejects a distorted picture.

## Caveats

- "Opens" means LibreOffice opens it; no PowerPoint in this environment. The rpptx comment anchor (`ac:spMk id`
  without `creationId`, P07) and the hand-written modern comments were not tried in PowerPoint.
- `builtin_validate` fails on every P07 output, rpptx's included: the skill's XSD set has no schema for the modern
  comment authors part (`p188:authorLst`: "No matching global declaration"). A validator false positive.
- pptxgenjs writes `[Content_Types]` overrides for `slideMaster2..10.xml`, which do not exist (P01 builtin,
  `oracle_info.package_warnings`); every validator here accepts it.
- Times are small and dominated by interpreter start (rpptx CLI about 0.02 s, the rpptx Python binding about 0.05
  s, python-pptx about 0.075 s to import); LibreOffice was warm.
- Builtin LOC includes the unzip and zip boilerplate (about 7 lines per script), as the skill's workflow requires.

## rdocx findings (pin 401e3f42, rpptx 0.14.4)

1. **A chart whose data labels have a thousands-separator number format renders as an "Unsupported chart"
   placeholder** (P12 fails on it; P02/P12/P13 deck, slide 6). `render`, `convert --to png/pdf` and the Python
   renders print "unsupported chart rendered as labelled placeholder: ChartML c:numFmt/@formatCode projection has
   invalid value: #,##0". `0`, `0.00` and `0.0%` draw; `#,##0`, `#,##0.0` and `"EUR "#,##0` do not. LibreOffice
   draws all of them. Reproduction (input with python-pptx, render with rpptx):

   ```python
   import subprocess, sys
   from pptx import Presentation
   from pptx.chart.data import CategoryChartData
   from pptx.enum.chart import XL_CHART_TYPE
   from pptx.util import Inches
   prs = Presentation()
   s = prs.slides.add_slide(prs.slide_layouts[6])
   d = CategoryChartData(); d.categories = ["Q1", "Q2"]; d.add_series("Revenue", (820, 910))
   plot = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(6), Inches(4), d).chart.plots[0]
   plot.has_data_labels = True
   plot.data_labels.number_format = "#,##0"; plot.data_labels.number_format_is_linked = False
   prs.save("nf.pptx")
   r = subprocess.run([sys.argv[1], "render", "nf.pptx", "-o", "nf", "--dpi", "30"], capture_output=True, text=True)
   print("unsupported chart" in r.stderr + r.stdout)   # True; expected False (format "0" gives False)
   ```

2. **The second content placeholder of the default "Two Content" layout renders 2.1 inches low and grey.** Its
   geometry (`effective_geometry()`, `text_layout()`) is right (y = 1.83 in, 28 pt), the render is not; seen in P01
   and P02 (rpptx's own default template and python-pptx's). LibreOffice draws it level with the left column.

   ```python
   import rpptx
   prs = rpptx.Presentation()
   s = prs.slides.add_slide(next(l for l in prs.slide_layouts if l.name == "Two Content"))
   s.shapes.title.text = "Two columns"
   s.placeholders[1].text = "Left column text"; s.placeholders[2].text = "Right column text"
   open("two.png", "wb").write(prs.render_slide_to_png(0, 50))
   # first ink row of the left half: 2.14 in, black; of the right half: 4.24 in, darkest grey 137.
   # Expected: both about 2.1 in and black.
   ```

3. **`fit_report()` / `rpptx fit` flag a table cell taller than its stored row height**, with `needed_scale` 40 %,
   although PowerPoint and rpptx's own renderer grow the row (python-api.md says so for `rows[k].height`) and the
   grown table stays on the slide (P06 false alarm). It does not tell this case from a table that runs off the
   slide. Also: the plain-text `rpptx fit` output prints the frame path as a Rust debug value (`Array
   [Number(1)]`) with no shape name, and lists every layout and master frame once per slide that uses it (55
   frames for a 3-slide deck); `--json` has the names.

   ```python
   import rpptx
   from rpptx.util import Inches
   prs = rpptx.Presentation()
   s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = "Risks"
   t = s.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(10), Inches(0.8)).table
   for r in range(2): t.rows[r].height = Inches(0.4)
   t.cell(0, 0).text, t.cell(0, 1).text, t.cell(1, 0).text = "Risk", "Mitigation", "Hiring"
   t.cell(1, 1).text = "Start interviews in November with two agencies; " * 4
   f = [x for x in prs.fit_report().frames if x.update.cell == (1, 1)][0]
   print(f.overflow, f.needed_scale)   # True 40000; the render ends the grown table at 3.9 in on a 7.5 in slide
   ```

4. **The Markdown and HTML export leave every chart out** ("slide 1/shape 1: unsupported chart, SmartArt or
   embedded-object projection omitted"), even a chart rpptx made itself; the built-in `slide_text.py` prints the
   chart's data as a table. The documented `&#32;` for every space of titles, runs and notes, and the escaped
   punctuation, make the file hard to read raw (P13 passes once decoded).

   ```python
   import rpptx
   from rpptx.util import Inches
   prs = rpptx.Presentation()
   s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = "Revenue"
   d = rpptx.CategoryChartData(); d.categories = ["Q1", "Q2"]; d.add_series("North", (1150, 1190))
   s.shapes.add_chart(rpptx.XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), d)
   print("1150" in prs.export_markdown().text)   # False; expected the chart's data (or at least its title)
   ```

5. **`prs.slides.index(slide)` (python-pptx's `Slides.index`) does not exist** and raises a bare `AttributeError`,
   not the error naming the rpptx way that python-api.md promises for missing python-pptx names; P04 writes
   `list(prs.slides).index(slide)`. With `slides.move(from_, to)` taking the final index, the first draft of P04
   moved the copy one slide too far (an agent-level slip, not a bug; handles that follow their slides made the
   rest short).
6. **Clumsy, not wrong (P10):** `table.rows.add_row()` gives the new cells no run (only `a:endParaRPr`), so the
   run-level edit SKILL.md recommends for keeping formatting (`paragraphs[0].runs[0].text = ...`) raises
   `IndexError`; `cell.text = ...` works and keeps the size through `endParaRPr`.

No rpptx output failed `rpptx validate`, the built-in `validate.py` (apart from the P07 schema gap above) or
LibreOffice.
