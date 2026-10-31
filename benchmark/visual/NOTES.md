# Visual: notes

Question: how does an agent get visual feedback on a Word or PowerPoint file, and how faithful is that picture
compared with what Word and PowerPoint print?

```bash
.cache/venv/bin/python visual/run.py            # about 2 minutes on 32 CPUs; results/visual.json, visual/figures/
.cache/venv/bin/python visual/run.py --only figures   # one part again, the others from .cache/work/visual/partial.json
```

| File | What |
|---|---|
| `run.py` | runner: the four parts below, then `results/visual.json` |
| `fidelity.py` | part 1: render each real file that has an Office-made PDF twin, compare page by page |
| `metrics.py` | Poppler readers (pages, words, fonts), SSIM, ink F1, word matching |
| `capabilities.py` | part 2: the feedback loop of each stack, each row tested |
| `findings.py` | the rdocx / rpptx faults found here, each as a neutral reproduction, re-run on the pin |
| `figures.py` | part 3: labelled side-by-side PNGs in `figures/` |

Versions: rdocx 0.16.4 / rpptx 0.14.4 (pin 401e3f42), LibreOffice 26.2.6.3, Poppler 26.01, Linux x86_64.

## Part 1: fidelity against Office's own PDF

Inputs: the 13 real files of `corpus/manifest.json` with a PDF twin (8 Word reports and forms, 5 decks). The
twin is the ground truth: Word or PowerPoint laid the pages out. 9 twins were saved by Word or PowerPoint itself
(`office-exported`: 8 docx, 1 deck), 4 decks went through another PDF writer (2 Mac PowerPoint via macOS Quartz,
2 Acrobat PDFMaker; `other-producer`): PowerPoint still laid them out.

Routes, each the shortest documented one:

- rdocx: `rdocx convert F --to pdf --font-dir .cache/fonts` (`rpptx convert` for decks). The PNG route
  (`rdocx render --dpi 60 --font-dir`) was compared as well: its SSIM against Office is within 0.03 of the PDF
  route on every document (`stacks.rdocx.png_route`), so the PDF route stands for what the agent sees.
- built-in: the skills' `scripts/office/soffice.py --headless --convert-to pdf` with `FONTCONFIG_FILE` (Carlito,
  Caladea and the system fonts).
- All three PDFs (Office, rdocx, LibreOffice) rasterised by the same pdftoppm, grayscale, 60 dpi.

Metrics (per document, then median and mean over documents; each document counts once):

- Page count against the Office PDF; pages paired by index (`pages_paired` = the smaller count).
- SSIM (skimage, Gaussian window sigma 1.5) of paired pages. SSIM punishes ink in the wrong place more than no
  ink: a white page scores 0.50 to 0.78 against the Office pages (`ssim_blank_page`), and a page whose text sits a
  few points lower can score below that floor. Read SSIM against the blank floor of the same document.
- Ink F1: share of ink pixels with ink of the other page within 3 px (3.6 pt), both ways, F1. Tolerant to small
  shifts, 0 for a white page.
- Words: every word box of the Office PDF (`pdftotext -bbox`) matched to the renderer's (difflib on normalised
  tokens over the whole document, so a word that moved to another page is still paired). `same_page`: share of
  Office words found on the same page; `within_5pt` / `within_20pt`: same page and left edge and vertical centre
  within 5 / 20 points; `text_pages_aligned`: pages (5 words or more) with 90% of their words on the same page,
  i.e. where the page breaks agree. Also the median signed offset (`median_dx_pt`, `median_dy_pt`).
- Fonts: share of the Office PDF's letters and digits per font family (`pdftohtml -xml`); a family is missing for
  a renderer when its PDF (`pdffonts`) has neither the family nor a metric-compatible clone (Carlito, Caladea,
  Liberation / Arimo / Tinos / Cousine).

### Per document

SSIM, ink F1 and words: rdocx / LibreOffice. Pages: Office / rdocx / LibreOffice.

| document | PDF | pages | SSIM (blank) | ink F1 | words same page | within 5 pt | pages aligned |
|---|---|---|---|---|---|---|---|
| charity-commission examiner template | Word | 2 / 2 / 2 | 0.69 / 0.74 (0.72) | 0.62 / 0.83 | 0.95 / 1.00 | 0.04 / 0.17 | 2/2, 2/2 |
| dbt flexible-working form | Word | 3 / 3 / 3 | 0.96 / 0.97 (0.78) | 0.99 / 1.00 | 1.00 / 1.00 | 0.93 / 1.00 | 3/3, 3/3 |
| pins local-plan Q&A (Aptos) | Word | 11 / 12 / 11 | 0.58 / 0.54 (0.63) | 0.65 / 0.73 | 0.52 / 0.81 | 0.06 / 0.07 | 0/11, 5/11 |
| dcms creative-industries consultation | Word | 22 / 21 / 22 | 0.59 / 0.60 (0.68) | 0.60 / 0.55 | 0.97 / 0.15 | 0.16 / 0.12 | 19/21, 5/21 |
| dcms taylor past schemes | Word | 47 / 48 / 47 | 0.51 / 0.86 (0.62) | 0.56 / 0.98 | 0.39 / 0.94 | 0.06 / 0.60 | 9/47, 37/47 |
| dcms taylor interim evaluation | Word | 66 / 67 / 66 | 0.52 / 0.90 (0.65) | 0.43 / 1.00 | 0.10 / 0.99 | 0.01 / 0.95 | 3/66, 65/66 |
| dcms social impacts | Word | 136 / refused / 137 | - / 0.45 (0.57) | - / 0.53 | - / 0.15 | - / 0.08 | -, 14/136 |
| dcms social impacts appendices | Word | 153 / refused / 153 | - / 0.85 (0.61) | - / 0.93 | - / 0.85 | - / 0.51 | -, 100/151 |
| ofqual regulator view (deck) | Mac Quartz | 13 / 13 / 13 | 0.95 / 0.95 (0.74) | 0.99 / 1.00 | 0.89 / 1.00 | 0.85 / 0.95 | 12/13, 13/13 |
| ofqual reform teachers (deck) | Mac Quartz | 28 / 28 / 28 | 0.92 / 0.88 (0.68) | 1.00 / 1.00 | 0.99 / 0.99 | 0.52 / 0.64 | 22/24, 22/24 |
| ofsted childminder (deck, Tahoma) | PDFMaker | 41 / 41 / 41 | 0.79 / 0.79 (0.72) | 0.96 / 0.92 | 0.95 / 0.99 | 0.68 / 0.24 | 33/34, 33/34 |
| dcms taylor technical annex (deck) | PowerPoint | 52 / 52 / 52 | 0.68 / 0.82 (0.50) | 0.93 / 1.00 | 0.93 / 0.95 | 0.77 / 0.89 | 41/52, 46/52 |
| ukces skills survey (deck, 64 charts) | PDFMaker | 69 / 69 / 69 | 0.63 / 0.84 (0.64) | 0.42 / 0.99 | 0.61 / 0.89 | 0.25 / 0.54 | 19/67, 43/67 |

### Aggregated (median / mean over documents)

| group | renderer | docs | exact page count | SSIM | ink F1 | words same page | within 5 pt | pages aligned |
|---|---|---|---|---|---|---|---|---|
| docx, Word-exported | rdocx | 6 (2 refused) | 2 of 6 | 0.59 / 0.64 | 0.61 / 0.64 | 0.74 / 0.66 | 0.06 / 0.21 | 0.55 / 0.52 |
| | LibreOffice, same 6 | 6 | 6 of 6 | 0.80 / 0.77 | 0.91 / 0.85 | 0.97 / 0.82 | 0.39 / 0.48 | 0.89 / 0.74 |
| | LibreOffice, all 8 | 8 | 7 of 8 | 0.80 / 0.74 | 0.88 / 0.82 | 0.90 / 0.74 | 0.34 / 0.44 | 0.72 / 0.65 |
| pptx, PowerPoint-exported | rpptx | 1 | 1 of 1 | 0.68 | 0.93 | 0.93 | 0.77 | 0.79 |
| | LibreOffice | 1 | 1 of 1 | 0.82 | 1.00 | 0.95 | 0.89 | 0.88 |
| pptx, other producer | rpptx | 4 | 4 of 4 | 0.86 / 0.82 | 0.98 / 0.84 | 0.92 / 0.86 | 0.60 / 0.57 | 0.92 / 0.77 |
| | LibreOffice | 4 | 4 of 4 | 0.86 / 0.86 | 0.99 / 0.98 | 0.99 / 0.97 | 0.59 / 0.59 | 0.94 / 0.88 |
| all 5 pptx | rpptx | 5 | 5 of 5 | 0.79 / 0.79 | 0.96 / 0.86 | 0.93 / 0.87 | 0.68 / 0.61 | 0.92 / 0.78 |
| | LibreOffice | 5 | 5 of 5 | 0.84 / 0.86 | 1.00 / 0.98 | 0.99 / 0.96 | 0.64 / 0.65 | 0.92 / 0.88 |

Reading:

- On Word documents LibreOffice is closer to Word than rdocx: same page count on 7 of 8 (rdocx 2 of 6), and on
  the two Taylor reports it keeps 94-99% of the words on Word's page where rdocx keeps 10-39%. rdocx refuses two
  of the eight files outright (duplicate bookmark ids, F11). rdocx wins one document clearly: the creative
  industries consultation, where LibreOffice runs a page ahead from page 6 (15% of words on Word's page, rdocx 97%;
  figure 4).
- rdocx's drift comes from a few layout differences that add up (F4 to F7 below, a header-to-body gap 7 pt larger
  than Word's on the Taylor reports, rows not split, frames placed differently), not from line spacing: on Taylor
  interim page 7 rdocx, LibreOffice and Word all set the body at a 13.9 pt pitch.
- On decks the page count is never in question (one slide, one page) and both renderers keep text on its slide;
  LibreOffice is closer on ink (pictures and charts: F2, F3, F8, F9 cost rpptx most on the UKCES and Taylor decks)
  and on text position for four of the five decks; rpptx is closer on the Ofsted deck (0.68 against 0.24 of the
  words within 5 pt: LibreOffice sets Tahoma's substitute 13 pt to the right).
- PDF producer: LibreOffice reaches an ink F1 of 0.99 to 1.00 against the PowerPoint-exported deck and against
  three of the four decks another program wrote (Quartz, PDFMaker), so the PDF writer does not seem to limit the
  comparison; the deck groups are too small to say more.

### Fonts

The Office PDFs draw almost everything in Arial, which both renderers replace with a metric clone; Calibri is
0 to 7% of the letters (Carlito covers it). Missing for both renderers:

| document | missing families (share of the letters) | dominates |
|---|---|---|
| ofsted childminder briefing | Tahoma (99.8%) | yes |
| pins local-plan Q&A | Aptos Narrow, Aptos (5.1%: headings) | no |
| ukces skills survey | Impact (0.3%) | no |

Calibri Light, Garamond, Open Sans and the other families the files declare are not drawn in the Office PDFs
(the manifest lists declared fonts; this table lists drawn ones). rdocx says when it substitutes (`Font 'Aptos' is
approximate: selected 'Carlito'`); LibreOffice says nothing.

## Part 2: the feedback loop

Every row ran here (`capabilities` in the JSON has the command and the evidence of each).

| capability | rdocx / rpptx | built-in (LibreOffice + pdftoppm) | pylib |
|---|---|---|---|
| one page or slide to PNG | pass: `rdocx render F --pages 1`, `rpptx render F --slide 1`, one command | pass: soffice.py converts the whole file to PDF, then `pdftoppm -f 1 -l 1` | unsupported (no renderer) |
| only selected pages | pass: `--pages 3,5`; decks `rpptx convert --to png --slides 3,5` | pass, but the whole file is converted first | unsupported |
| tracked changes in the render (docx-M) | pass, opt-in: `--revision-view tracked [--revision-style legal]`: 32 of 33 deletions shown; the default render shows the accepted text | pass, by default: 31 of 33 shown with LibreOffice's marks | unsupported |
| comments in the render | unsupported (no option; `rdocx comment list --json` gives text and anchor) | partial: 0 of 25 by default; 25 of 25 with an export option the skill does not document (`ExportNotesInMargin`) | unsupported |
| thumbnail grid, decks | pass: `render_all_handouts(9)` (Python; handout pages, slide numbers, no file names) | pass: `scripts/thumbnail.py`, labelled grids | unsupported |
| thumbnail grid, documents | unsupported | unsupported (`thumbnail.py` refuses .docx) | unsupported |
| page of each heading / block | pass: `rdocx layout --json` gives every block's page; Word's page for 6 of 13 headings of the two Taylor reports | partial: no layout API; LibreOffice's PDF bookmarks (not in the skill) give Word's page for 11 of 13 | partial: python-docx `rendered_page_breaks` (Word's own breaks at its last save): 7 of 13; none in a file Word did not save, stale after an edit |
| text-overflow report, slides | pass: `rpptx fit --json` flags the overflowing frame of a fixture and not the other; exits 2 on pptx-M and pptx-L (duplicate shape id, F10) | unsupported: the skill says to look at the images | partial: `TextFrame.fit_text(font_file=...)` on a copy flags the fixture correctly, rewriting the size |
| warns about a substituted font | pass (stderr, `*_with_diagnostics`) | unsupported (silent) | unsupported |
| time to first picture of page 1, one run | docx-M 0.8 s, docx-L 6.4 s, pptx-M 0.65 s, pptx-L 1.7 s | docx-M 1.1 s, docx-L 3.9 s, pptx-M 1.6 s, pptx-L 3.5 s | - |

Notes: rdocx lays the whole document out before drawing any page, so page 1 of docx-L (195 pages) costs as much
as `rdocx layout` of the whole file (about 7 s, measured once by hand), more than LibreOffice's full conversion
(about 4 s). Times are one run each, after other runs of the same session (warm file caches), and move by about
15% between runs (docx-L: 6.4 to 7.3 s over three full runs); the speed section times properly.

## Part 3: figures

| file | what |
|---|---|
| `figures/fig1_docx_rdocx_closest.png` | the Word page where rdocx scores its best SSIM (DBT flexible-working form, page 2): all three agree |
| `figures/fig2_docx_rdocx_farthest.png` | rdocx's lowest SSIM (creative industries, page 12): the words are on the right page, the table is drawn wider and the box borders differ; SSIM ranks this below LibreOffice's page, which shows other text |
| `figures/fig3_docx_libreoffice_closest.png` | LibreOffice's best (Taylor past schemes, page 3, the contents): rdocx renumbers the contents with its own pages, underlines the entries and draws the footer table's borders (F5) |
| `figures/fig4_docx_libreoffice_farthest.png` | LibreOffice's worst (creative industries, page 14): LibreOffice runs a page ahead, rdocx matches Word |
| `figures/fig5_pptx_slide_median.png` | the slide at rpptx's median SSIM (Taylor technical annex, slide 13, PowerPoint's PDF): rpptx draws the master's hidden guide grid (F8) |
| `figures/fig6_tracked_changes.png` | one replacement in docx-M: rdocx tracked view, rdocx legal style, LibreOffice's default export |
| `figures/fig7_thumbnail_grids.png` | pptx-M: built-in `thumbnail.py` grid and rpptx's 9-up handout; the handout shows rpptx's chart placeholders (F9) |

Pages were picked by rule (`figures.py`): pages of the Word reports with 80 words or more, not page 1, no picture
in the Office PDF (no logos or third-party images), highest and lowest SSIM per renderer; the slide is the
median-SSIM slide without pictures. Figures 6 and 7 use synthetic files.

Attribution: figures 1 to 5 contain public sector information licensed under the Open Government Licence v3.0:
Department for Business and Trade, "Request flexible working: application form template"
(https://www.gov.uk/government/publications/the-right-to-request-flexible-working-form); Department for Culture,
Media and Sport, "Classifying and Measuring the Creative Industries: consultation"
(https://www.gov.uk/government/consultations/classifying-and-measuring-the-creative-industries-consultation-on-proposed-changes);
Department for Culture, Media and Sport (by Frontier Economics), "Evaluation of the Taylor Review pilot" report on
past schemes and technical annex
(https://www.gov.uk/government/publications/evaluation-of-the-taylor-review-pilot-emerging-monitoring-and-evaluation-evidence-september-2018-to-march-2019).

## Judgement calls

- The PDF route for rdocx/rpptx, not `render`: the same rasteriser for the three PDFs isolates layout from
  anti-aliasing, and word boxes need a PDF. The PNG route is reported beside it (within 0.03 SSIM).
- rpptx's PDFs put each run's whole text in the `/ActualText` of its first word (F1), so `pdftotext -bbox` gives one
  box per run. For the word metric only, rpptx PDFs are re-written by `pdftocairo -pdf` (glyphs and ToUnicode kept,
  ActualText dropped). The same round trip on LibreOffice's deck PDFs changes their word scores by 0.01 at most
  (`words_cairo_roundtrip_check`). Without it rpptx would score 1 to 18% "matched".
- Pages are paired by index, as the task asks: after a page break differs, every later page compares different
  text. `same_page` and `text_pages_aligned` say how much of that there is; ink F1 and SSIM then measure it.
- "Page of each heading" is scored against Word's PDF outline (bookmarks) of the two Taylor reports, on the 13
  headings found exactly once among the document's paragraphs (repeated titles left out).
- The built-in route for heading pages (PDF bookmarks through `pdftohtml -xml`) and for comments
  (`ExportNotesInMargin`) is not in the skill; both are recorded as `partial` with that said.
- `pylib` has no renderer; its rows are `unsupported` with the probe of python-docx's and python-pptx's public names,
  except where the library has a documented way (`rendered_page_breaks`, `fit_text`).

## Caveats

- 13 documents; the deck groups are small (1 PowerPoint-exported, 4 others). Office PDFs come from several Word and
  PowerPoint versions (2010 to 365, Mac); none was made on this machine.
- SSIM is a poor ranking of layout fidelity on its own (figure 2): read it with the blank floor, ink F1 and words.
- LibreOffice also uses the system fonts here (Noto, DejaVu...), rdocx's CLI its bundled ones and `--font-dir`;
  rdocx's Python binding falls back to system fonts too (F12), so CLI and Python can paginate differently.

## rdocx findings

Each is a function of `findings.py` that builds its own input (python-docx, python-pptx, pptxgenjs, zlib) and
records `reproduced` in `results/visual.json` (`rdocx_findings`); all 12 reproduce on pin 401e3f42. Shortest form:

1. **rpptx PDF text: a run's whole text is the ActualText of its first word.** `rpptx convert` of a text box "Alpha
   beta gamma" (python-pptx `add_textbox(...).text_frame.text = "Alpha beta gamma"`): `pdftotext -bbox` gives one
   word "Alpha beta gamma" with Alpha's box; expected three boxes. Search highlighting, text selection and any
   word-position tool see the line on its first word. rdocx's PDFs do not do this.
2. **rpptx does not draw an interlaced (Adam7) PNG** and reports `missing slide picture image relationship rId3`
   (the relationship and the part are there). Repro: a 64x48 red PNG written interlaced (zlib, 7 passes; Pillow
   cannot write one), added with python-pptx next to the same PNG not interlaced: the PDF holds 1 of 2 images.
   rdocx draws it. All five real decks have one or more (the Ofqual logo, 12 in the UKCES deck).
3. **EMF pictures are not drawn**: rdocx leaves an empty frame and says nothing; rpptx says `missing slide picture
   image relationship`. Repro: a PNG added with python-docx / python-pptx, its part swapped for an EMF that
   LibreOffice made from an SVG rectangle (`soffice.py --convert-to emf box.svg`), content type `image/x-emf`.
   LibreOffice draws it. Office files carry EMF logos and figures (the Taylor reports' cover logo, three in the
   Taylor annex deck).
4. **A Wingdings 2 bullet (U+F0A1 in `w:lvlText`) draws a missing-glyph box in Carlito, with no diagnostic.** The
   Wingdings (U+F0A7) and Symbol (U+F0B7) bullets of the same repro map to U+25AA and U+2022. LibreOffice draws a
   black square. The Taylor reports use it for every bullet.
5. **A table style's borders are drawn although the table's `w:tblBorders` set every border to none.** Repro:
   python-docx table with style "Table Grid", then `<w:tblBorders>` (top ... insideV all `none`) before
   `w:tblLook`: rdocx draws three horizontal lines, LibreOffice none (direct table properties win over the style).
   Seen on the Taylor reports' footers and boxes.
6. **A table row whose cells are vertically centred is moved whole to the next page instead of split.** Repro: 15
   short paragraphs, then a one-cell table of 40 lines with `<w:vAlign w:val="center"/>`, no `w:cantSplit`:
   0 of its 40 lines on page 1 (LibreOffice 22); without vAlign rdocx splits it (21 lines on page 1). The Local
   Plan Q&A's centred rows move whole from page 1 on (0 of 11 pages aligned with Word, LibreOffice 5).
7. **Automatic text colour on a black cell stays black.** Repro: a cell with `<w:shd w:val="clear" w:color="auto"
   w:fill="000000"/>` and default text: the PDF text is #000000; LibreOffice (and Word) draw it white. The charity
   template's section headings vanish.
8. **rpptx draws shapes and groups marked hidden** (`p:cNvPr hidden="1"`): a hidden rectangle, a hidden group and a
   visible rectangle (python-pptx, attribute set on `cNvPr`): rpptx draws all three, LibreOffice one. The Taylor
   annex master hides a dashed guide grid (figure 5).
9. **rpptx draws pptxgenjs bar and line charts as "Unsupported chart" placeholders**: `ChartML contains duplicate
   c:axId`, while the three ids are distinct (pptxgenjs writes a third `c:axId` for a series axis it does not
   declare). Repro: `p.addSlide().addChart(p.charts.BAR, [{name, labels, values}], {x, y, w, h})` and the same with
   `LINE`. 11 of pptx-M's 16 charts; LibreOffice draws them (figure 7).
10. **`rpptx fit` exits 2 on a deck with a duplicate shape id** (`autofit ...: duplicate shape ID 2`), which
    `rpptx render` accepts and PowerPoint repairs; pptxgenjs writes one on table slides, so `fit` gives no report
    for pptx-M and pptx-L. Repro: two python-pptx text boxes, the second's `cNvPr id` set to the first's.
11. **A document whose bookmarks share a `w:id` is refused** by render, text, inspect and validate (already in
    `corpus/NOTES.md`, repro there): 2 of the 8 Word twins cannot be rendered at all.
12. **The CLI and the Python binding draw the same document with different fonts.** `rdocx convert --font-dir DIR`
    draws the bullets of finding 4 in Carlito; `Document.to_pdf(font_dir=DIR)` draws U+F0A1 with FontAwesome (a
    system icon font: a "bullhorn") and U+F0B7's bullet with DejaVu Sans. On the Taylor past-schemes report the CLI
    gives 48 pages, Python 47 (checked once by hand); `page_count()` without `font_dir` says 48 too.

Observed, not reduced: rdocx rewrites a table of contents' page numbers with its own pagination and underlines its
entries (figure 3); framed paragraphs (`w:framePr`, `yAlign="bottom"`) land elsewhere than in Word (Taylor reports,
page 2); combo (bar and line) charts of the UKCES deck are placeholders (rpptx says so).
