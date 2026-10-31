# Corpus: notes

The inputs of the speed and visual sections: synthetic documents and decks in three sizes, built here from
committed generators with no network, and public Office files pinned by URL and SHA-256 in `real.json` (fetched,
never committed).

```bash
.cache/venv/bin/python corpus/build.py          # synthetic, ~15 s with the checks (32 CPUs)
.cache/venv/bin/python corpus/build.py --real   # download real.json into .cache/corpus/real/, check
.cache/venv/bin/python corpus/build.py --all    # both; --no-checks skips the open/validate checks
```

`build.py` is this section's runner. It writes `corpus/manifest.json`; a run of one half keeps the other half's
entries.

## Files

| Path | What |
|---|---|
| `build.py` | builds, checks, writes the manifest |
| `content.py` | the text and structure of every synthetic file (plans), seeded per file |
| `images.py` | the 48 PNG pictures (Pillow), seeded |
| `gen_docx.js` | docx-js generator (plan to .docx) |
| `gen_pptx.js` | pptxgenjs generator (plan to .pptx), then the built-in skill's `apply_theme.js` |
| `gen_pptx_pythonpptx.py` | python-pptx generator (same deck plan) |
| `deterministic.js` | freezes the clock and seeds `Math.random` / `crypto.getRandomValues` in the node generators |
| `ooxml_stats.py` | neutral counts (zipfile + lxml), self-contained so that it runs with `python -I` |
| `real.json` | the public files: URL, SHA-256, licence, producer, pages, PDF twin |
| `manifest.json` | written by `build.py`: every file, its counts, pages, checks |

Output, all under `.cache/corpus/`: `docx/`, `pptx/`, `plans/` (the JSON each file was built from, useful to
pick texts for find-and-replace tasks), `assets/` (pictures), `lo_pdf/` (LibreOffice's PDF of every file, from
the checks), `real/<id>/office/` and `real/<id>/pdf/` (one folder per download).

## Method

Content is the annual report (and the board briefing deck) of a fictional regional water authority: chapters
of prose built from ~25 sentence templates filled with districts, programmes, money, percentages, dates and
counts, so that sentence length varies from 3 to 45 words and numbers are everywhere. Paragraph text is split
into 1 to 7 runs at random word boundaries, about 15% of them bold or italic, and now and then a word is split
inside an unformatted run, as Word's edit history (rsid splits) leaves text.

docx (docx-js 9.15.0, the built-in docx skill's generator, so the corpus is not tuned to rdocx), following that
skill's gotchas: A4, `HeadingLevel` 1-3 (Cambria) over a Calibri 11 body, bullet and numbered lists from a
`numbering` config (each numbered list its own instance, so it restarts), tables with `columnWidths` and a DXA
width on every cell, `ShadingType.CLEAR` header and zebra shading, two-row headers with merged cells (`rowSpan`,
`columnSpan`) in the finance tables (10 to 18 body rows), PNG figures with captions, footnotes, a TOC field over
headings 1-3 (docx-js writes its entries without page numbers, dirty, so Word offers to update it), a header and
a "Page X of Y" footer (PAGE / NUMPAGES), a second section in landscape (Annex A, wide 10-column tables) that
inherits the header and footer. Comments: one per two target pages; tracked changes: one per target page,
cycling insertion, deletion and replacement (deletion plus insertion), four fictional authors, fixed dates.
Footnotes: one per three target pages.

| id | target | LibreOffice pages | rdocx pages |
|---|---|---|---|
| docx-S | 5 | 6 | 6 |
| docx-M | 50 | 52 | 51 |
| docx-L | 200 | 198 | 195 |

docx-M-clean: 51 / 51, docx-M-lo: 52 / 51 (the manifest has the numbers of the last build). `docx-M-clean` is docx-M's plan without comments and
tracked changes. `docx-M-lo` is docx-M opened and saved by LibreOffice (`soffice.py --headless --convert-to
docx`), a second producer's XML: LibreOffice's styles, numbering, settings, five header and five footer parts,
`LibreOffice/...` in app.xml.

pptx (pptxgenjs 4.0.1) as the built-in pptx skill's structured deck: a `THEME` object (Cambria headings,
Calibri body, a teal palette), `LAYOUT_WIDE`, seven layouts from `defineSlideMaster` (title, section divider,
title only, title and body, body and picture, two columns, chart and commentary) with named placeholders, footer
text and slide number on the layouts, `pres.SchemeColor` everywhere except the hex-only options, one section per
chapter (`addSection` + `sectionTitle`), speaker notes on every slide, native bar / line / pie charts with
titles, data labels, theme-coloured series and chart fonts `+mn-lt`, tables up to 8 x 6 with a shaded header,
pictures, stat callouts (rounded rectangles with 60 pt numbers), long-text slides (three paragraphs at 14 pt),
then `writeFile()` and the skill's `scripts/apply_theme.js`. 15 / 60 / 150 slides. `pptx-M-pythonpptx` builds
the same 60-slide plan with python-pptx 1.0.2 on its default template (4:3, Office theme, its own layouts,
`add_chart`, `add_table`, `add_picture`, `add_shape`, notes); python-pptx has no section API, so it has none.

Checks on every file (each a separate process; failures are recorded in the manifest, never hidden): `rdocx
inspect --json` / `rpptx inspect --json`, rdocx `page_count(font_dir=.cache/fonts)`, `rdocx validate` /
`rpptx validate`, python-docx / python-pptx open, LibreOffice PDF export through the built-in skill's
`scripts/office/soffice.py` (with the benchmark's `FONTCONFIG_FILE`), and the built-in skill's
`scripts/office/validate.py` (without `--original`).

## Judgement calls

- One generator per stack's documented creation path: docx-js for Word (the built-in docx skill's only
  creation path) and pptxgenjs for PowerPoint, plus python-docx's sibling python-pptx as the second deck
  producer and LibreOffice as the second Word producer. No rdocx-written file is in the corpus.
- docx-js 9.15 can estimate TOC page numbers (`pageNumbers: estimatePageNumbers` from `docx/layout`) and
  write native charts (`docx/charts`); the built-in skill teaches neither, so the corpus uses neither: the TOC
  has entries but no page numbers, and Word pictures are PNG figures.
- The docx header and footer are set on the first section only; the landscape section inherits them (no
  `headerReference`), as Word does when "Link to previous" is on.
- Reproducibility: node runs with `TZ=UTC`, `deterministic.js` fixes `Date` (core.xml dates, zip entry times)
  and seeds the random ids; python-pptx's and LibreOffice's zip entries get a fixed timestamp after saving
  (same names, order, data and compression; LibreOffice ignores `SOURCE_DATE_EPOCH` for OOXML), and python-pptx's
  embedded chart workbooks get a fixed `dcterms:created` (XlsxWriter stamps the current time). Two builds give the
  same SHA-256 for all nine files.
- Page targets are calibrated on LibreOffice's layout with Carlito and Caladea; Word will paginate differently.

## Caveats and odd things

- pptxgenjs writes `Microsoft Office PowerPoint` / `16.0000` in `docProps/app.xml`, and python-pptx's default
  template says `Microsoft Macintosh PowerPoint` / `14.0000`: app.xml `Application` does not prove a deck was
  made by PowerPoint. docx-js writes no `Application` at all. For the real corpus the PDF twin's producer and
  the file's own XML habits were checked as well (below).
- pptxgenjs writes duplicate shape ids on some table slides: the table's graphic frame can get `id="25"`, the id
  it always gives the slide-number placeholder (pptx-M and pptx-L, slide 24). `rpptx validate` reports
  `DuplicateShapeId { slide: 23, id: 25 }` (0-based slide index); the built-in `validate.py`, python-pptx and
  LibreOffice accept the deck. Kept as generated (it is what the built-in path produces) and recorded as a failed
  check.
- `rdocx inspect` counts 525 "paragraphs" in docx-M where the manifest counts 1,984 `w:p`: rdocx counts the
  accepted main body's top-level paragraphs, the manifest every `w:p` in document.xml.
- `docx-M-lo`: rdocx inspect reports `Unknown Symbol symbol U+F02D: original code retained` (LibreOffice
  rewrites the level-2 bullet as a Symbol-font character).

## Real corpus

14 public files (9 .docx, 5 .pptx), all but one from GOV.UK under the Open Government Licence v3.0 (the
publication pages state no exception), one French note from data.gouv.fr under Licence Ouverte 2.0. Each was
checked to be produced by Word or PowerPoint (app.xml `Application`, cross-checked with the PDF twin's Creator).
13 have a PDF twin from the same publication page with the same page count:

| id | kind | pages / slides | PDF twin |
|---|---|---|---|
| gov-uk-charity-commission-independent-examiner-template | docx | 2 | Word 2013, 2 |
| gov-uk-dbt-flexible-working-request-form | docx | 3 | Word for Microsoft 365, 3 |
| data-gouv-ssmsi-cvs-note-presentation | docx | 5 (French) | none |
| gov-uk-pins-local-plan-webinar-qa | docx | 11, A4 landscape, Aptos | Word for Microsoft 365, 11 |
| gov-uk-dcms-creative-industries-consultation | docx | 22 | Word 2010, 22 |
| gov-uk-dcms-taylor-review-past-schemes | docx | 47 | Word for Office 365, 47 |
| gov-uk-dcms-taylor-review-interim-evaluation | docx | 66, 11 charts | Word for Office 365, 66 |
| gov-uk-dcms-social-impacts-culture-sport | docx | 136, 17 MB | Word 2013, 136 |
| gov-uk-dcms-social-impacts-culture-sport-appendices | docx | 153, 9 charts, 14 landscape sections | Word 2013, 153 |
| gov-uk-ofqual-skills-employability-regulator-view | pptx | 13 | Mac PowerPoint via macOS Quartz, 13 |
| gov-uk-ofqual-qualifications-reform-teachers | pptx | 28, notes on every slide | Mac PowerPoint via macOS Quartz, 28 |
| gov-uk-ofsted-childminder-pre-registration-briefing | pptx | 41, 6.8 MB of pictures | Acrobat PDFMaker in PowerPoint, 41 |
| gov-uk-dcms-taylor-review-technical-annex-deck | pptx | 52, A4-landscape slides, text-dense | PowerPoint for Office 365, 52 |
| gov-uk-ukces-employer-skills-survey-2015-slide-pack | pptx | 69, 64 charts | Acrobat PDFMaker in PowerPoint, 69 |

`route` in `real.json` says how each PDF was made: `office-native` (Word or PowerPoint "Save as PDF" on Windows;
the 9 best ground truths: 8 docx and 1 deck), `office-mac-quartz` (Office for Mac lays the pages out, macOS writes
the PDF), `acrobat-pdfmaker` (PowerPoint lays out, Adobe's add-in writes the PDF). For the visual section, the
native twins are the cleanest; the others are PowerPoint's layout through another PDF writer.

Method: a helper agent searched GOV.UK's content and search APIs, data.gouv.fr and US/EU sites for about 20
minutes, downloaded 60 candidates (each into its own folder, inspected only with `python -I` scripts and
`pdfinfo`) and kept 17; I dropped three of them and re-fetched the 14 through `build.py --real`, which verified
every SHA-256. Text was scanned (zipfile + lxml, `python -I`) for copyright lines and personal data.

Judgement calls:

- Dropped `School Games 2012-13 statistical release` (the only real file with tracked changes, 2 leftover
  insertions) because it names the responsible statistician with a direct phone number; dropped the `Tech City UK`
  evaluation (contractor staff e-mail and phone, and its Mac/Quartz PDF has 129 pages against app.xml's 62);
  dropped the PHE tuberculosis slideset (an Ordnance Survey map under its own Crown copyright and database
  right, and no PDF twin). Kept: generic enquiry lines (creative industries consultation, Ofqual, UKCES) and the
  academic authors' byline in the Social Impacts appendices.
- Four files were written for a department by contractors (Frontier Economics: the three Taylor Review files;
  IFF Research: the UKCES survey). GOV.UK publishes them under the OGL and no third-party copyright line was
  found in their text.
- Bands by the twin's page count: short <= 10, medium 11-80, long > 80. The short band has two pairs (2 and 3
  pages) and the French note without a twin; GOV.UK seldom publishes short documents in both formats. No deck
  near 100 slides has a twin (largest: 69).
- The real files carry no comments and no tracked changes; the synthetic corpus covers both.
- Sources that did not work: curl was refused by most US federal sites tried (ed.gov, hhs.gov, cdc.gov,
  fema.gov, usda.gov, nih.gov, ahrq.gov) and by EUR-Lex; regulations.gov had decks as PDF only; EU candidates had
  PDFs printed by Chrome. Rejected pairs: street-works code of practice (PDF 158 pages vs 168, PDFMaker), a
  trustees' report template (PDF from InDesign), a Welsh examiner template and a broadband study (the PDF was
  saved by another Word version than the docx).
- `pages.app_xml` is Word's own count at its last save and can be stale: both Social Impacts files say 1 (their
  PDFs have 136 and 153 pages). `pdf_twin.pages_match` compares the twin with app.xml, LibreOffice and rdocx.
- Author names in `docProps/core.xml` are not read or recorded.

Checks on the real files (manifest `checks_failed`):

- On the first run, rdocx refused to open both Social Impacts files (`duplicate bookmark id`, see rdocx findings;
  fixed since); python-docx,
  LibreOffice and the built-in `validate.py` (main report) open them.
- The built-in `validate.py` (without `--original`) fails four untouched Office files: Word charts with
  `c:showDLblsOverMax` / an empty `c:ext` (Taylor interim evaluation, Social Impacts appendices) and PowerPoint's
  `ppt/revisionInfo.xml` / `ppt/authors.xml` (Ofsted, UKCES), parts its schema set does not declare. These are
  false positives on files Office wrote.
- Pages: LibreOffice matches the Office PDF on 7 of 8 docx twins (Social Impacts: 137 vs 136); rdocx on 3 of 6
  that it opens (creative industries 21 vs 22, Taylor interim 67 vs 66, Local Plan Q&A 12 vs 11, the last set in
  Aptos, which neither renderer has: rdocx reports `Font 'Aptos' is approximate: selected 'Carlito'`).

## rdocx findings

Met on the first run (2026-10-31, pin `401e3f42`). They were reported upstream as tensorbee/rdocx#431 to #437, and the pin the benchmark now runs on carries the fixes of all but the open gaps of `tests/gaps.py`. Kept as the record of what the benchmark found; the README has the current results.

1. **rdocx refuses a document whose bookmarks share a `w:id`** (Word 2013 saved two GOV.UK reports that way:
   `_GoBack` reuses the id of a `_Toc` bookmark in the body, or of a bookmark in a header). `rdocx text`,
   `inspect`, `validate` and `rdocx.Document()` all fail with `cannot scan identifiers in XML part
   /word/document.xml: duplicate bookmark id 2 in imported or preserved XML`; Word, LibreOffice and python-docx
   open the files. Expected: open (the ids only pair `bookmarkStart` with `bookmarkEnd`), at most a diagnostic,
   and renumber on save if needed. Neutral reproduction (pinned build 401e3f42, rdocx 0.16.4):

   ```python
   import docx
   from docx.oxml import parse_xml
   from docx.oxml.ns import nsdecls

   def bookmark(p, bid, name):
       p._p.insert(0, parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bid}" w:name="{name}"/>'))
       p._p.append(parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bid}"/>'))

   d = docx.Document()
   bookmark(d.add_paragraph("First heading"), 2, "_Toc378755770")
   bookmark(d.add_paragraph("Some text"), 2, "_GoBack")
   d.save("same-part.docx")
   # $R/rdocx text same-part.docx -> Error: cannot scan identifiers in XML part /word/document.xml:
   #   duplicate bookmark id 2 in imported or preserved XML (exit 1)
   # The same with one bookmark in the body and one in the header (same id) fails on /word/header1.xml.
   ```

2. Not a bug: `rpptx validate` reports pptxgenjs's duplicate shape ids (above), which the other checkers miss.
