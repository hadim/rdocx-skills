# Word authoring and editing tasks D01-D14: notes

Runner: `.cache/venv/bin/python features/docx/run_authoring.py` (all tasks, about 45 s) writes
`results/features_docx_authoring.json`; `--task D05 --quick` runs one task without timings or validators.

## Method

- **Inputs.** Each task's `make_input.py` builds its inputs once per run into `.cache/work/features_docx_authoring/Dxx/in/`,
  and the three stacks get the same files. Documents are built with python-docx (its default template, a Word 2011
  file) plus `authoring_inputs.py`, which makes them look like Word's output: paragraphs of several runs with different
  formatting, runs split mid-word with revision ids (`w:rsidR`) and `w:proofErr` marks between them, `_GoBack`
  bookmarks, `w:rsids` in the settings, real Word structures where a task needs them (DrawingML text box with its VML
  fallback, footnotes part with separators, comments with their styles, `List Paragraph` + `w:numPr` lists). Every
  input passes `rdocx validate` and the built-in `validate.py`.
- **Implementations.** `rdocx.py` (pinned Python, `docx_ops` on the path) or `rdocx.sh` (CLI); `builtin.sh` plus the
  scripts it calls; `pylib.py` (python-docx, lxml only where python-docx has no API). Arguments: inputs, then the
  output, then task extras (D03: the expected count).
- **Oracle.** `check.py OUTPUT` reads the output with zipfile + lxml only (`authoring_oracle.py`: style names
  through `styles.xml`, numbering formats through `numbering.xml`, inherited headers and footers, field
  instructions, per-run effective bold/italic). Lines starting `info:` describe the output and never fail it; they are
  copied into `notes`.
- **Validity.** On every .docx output: `rdocx validate`, the built-in `scripts/office/validate.py OUTPUT`
  (`--original INPUT` for edits), and a LibreOffice PDF conversion through `scripts/office/soffice.py`
  (`libreoffice_opens` = exit 0 and a PDF over 1 KB). JSON and Markdown outputs get `valid: null`.
- **Timing.** Wall time of the whole implementation as an agent runs it (interpreter start included), median of 5
  runs after a warm-up, one stack after the other (validators run afterwards, in parallel).
- **LOC.** Non-blank lines that are not `#` or `//` comments, over every file of the stack (`rdocx.*`, `builtin*`,
  `pylib.*`); not `make_input.py` nor `check.py`. Shell boilerplate (`set -euo pipefail`) counts.
- **processes.** 1 for a script; for a shell script, each command line (unzip, find, merge_runs.py, the edit script,
  zip). Variable assignments and `mktemp` do not count. Verification (validate, render) is excluded for every stack:
  the harness runs it uniformly.
- **raw_xml.** True when the implementation writes or edits OOXML by hand: XML tag strings, `qn()`, `OxmlElement`,
  python-docx private elements (`._p`, `._tr`...), `replace_xml`. Detected by pattern on the files and checked by
  hand. For JSON/Markdown outputs, parsing parts by hand is noted (`parses the OOXML parts by hand`) but not counted.
- **tools.** Detected from the files (imports, commands).

## Results

The tables of `benchmark/README.md`, generated from `results/features_docx_authoring.json` by `report.py`. This file does not repeat the numbers: they change with each run.

## Judgement calls

- **rdocx.py is run with `python -P`**: the task's file name `rdocx.py` would otherwise shadow the `rdocx` module
  (the script's folder comes first on `sys.path`). An agent would name its script otherwise; no effect on the code.
- **The built-in stack's edits** follow the SKILL.md workflow (`unzip`, `find -type l -delete`, `merge_runs.py`, edit
  `word/document.xml`, `zip -Xr`). An agent edits the XML with its editor; here the edit is a Python script with
  string or regex surgery (closest to the Edit tool) or lxml where the edit is structural (D02 read, D03, D05). Parts
  the skill has no helper for (headers, footers, footnotes, charts, relationships, content types) are written by hand.
- **D02 builtin reads the XML, not pandoc.** The skill's reading path is `pandoc -t markdown`; verified on the input,
  pandoc drops headers and footers, puts the text box's text in the body flow, and gives comments only with
  `--track-changes=all`, inline in the anchoring paragraph. An agent asked for text per story falls back to the parts.
- **D13 builtin uses `pandoc in.md -o out.docx`**, not docx-js: the skill documents pandoc for reading only, but it
  is one command with the stack's own tool; with docx-js the agent would hand-translate the content into a script.
  D14 builtin is the skill's own `pandoc -t markdown`.
- **D08**: builtin writes the chart part by hand and its embedded workbook with openpyxl (installed with the stack,
  the xlsx skill's library); pylib uses python-pptx's `ChartXmlWriter` and `xlsx_blob` (python-pptx is part of the
  pylib stack) and adds the parts with python-docx's OPC classes. The oracle asks for an embedded workbook
  (`c:externalData`), which is what makes Word's "Edit Data" work.
- **D04 template syntax** is neutral (`{{client}}`, an item row of `{{name}}`/`{{qty}}`/`{{price}}`, a total row), not
  rdocx's `render_template` grammar; every stack clones the row its own way.
- **D05**: rdocx merges with `set_cell_grid_span(4, 0, 2)`, the first merge call SKILL.md lists; `Cell.merge` leaves
  an empty paragraph (finding 5), which the oracle rejects because Word drops an empty cell's paragraph when it
  merges. The new 3 cm column makes the table 19 cm wide on a 16 cm text width in all three stacks (info line).
- **D10 and D13 ask that links look like links** (Hyperlink style or underline), as Word, pandoc and LibreOffice
  write them. rdocx writes no style (finding 3): both rdocx implementations colour and underline the link runs
  themselves, which an agent following rule 7 (render and look) would do; a plain `rdocx convert notes.md` fails D13.
- **D09 checks the OMML structure the schema requires** (`m:rad` = `m:radPr?, m:deg, m:e`): without `m:deg`,
  LibreOffice reads the radicand as the degree and draws empty boxes for the radicand and the denominator. rdocx's
  own render draws it right, and `rdocx validate` passes it; only the built-in XSD check reports it.
- **D11** is `pass` for rdocx: the footnote is placed and numbered, LibreOffice renders it; the XSD error (finding 2)
  shows in `valid`. Word was not available to see whether it repairs the part.
- **D01 pylib** writes the TOC as an empty complex field and sets `w:updateFields` by appending it to the settings
  (the common snippet); docx-js writes TOC entries without page numbers; rdocx `insert_toc` + `rebuild_toc` writes
  entries with page numbers from its layout.
- **D03** refusal: each implementation gets the expected count; with 4 instead of 5 it must exit non-zero and write
  no output. All three pass (rdocx `replace_batch`; builtin and pylib count first, then exit).

## Caveats

- No Word here: `libreoffice_opens` and the XSD check stand in for "Word opens it". The rendering checks above were
  done by eye on LibreOffice PDFs and rdocx renders.
- The inputs come from python-docx's default template, whose `styles.xml` is 349 KB (every built-in style defined).
  rdocx's per-edit cost follows that part (finding 7), so its edit timings here are higher than on a Word 365 file
  with a 30-40 KB styles part.
- Implementations are written once by the benchmark author for each stack; a different competent agent could write
  fewer or more lines. The oracle's info lines record behaviour that is not graded (TOC entries, separator notes,
  grid width, which restart method).

## rdocx findings

Met on the first run (2026-10-31, pin `401e3f42`). They were reported upstream as tensorbee/rdocx#431 to #437, and the pin the benchmark now runs on carries the fixes of all but the open gaps of `tests/gaps.py`. Kept as the record of what the benchmark found; the README has the current results.

1. **`add_equation` writes `m:rad` without `m:deg`** for a square root (LaTeX `\sqrt{...}` and MathML `<msqrt>`).
   ECMA-376 requires `m:deg` (empty, with `m:degHide`). LibreOffice then shows empty boxes; `rdocx validate` passes,
   the built-in `validate.py` reports "Element m:e: This element is not expected. Expected is (m:deg)". An n-th root
   (`\sqrt[3]{x}`) is correct.
   ```python
   import re, rdocx
   d = rdocx.Document(); p = d.add_paragraph("")
   p.add_equation(r"\sqrt{x}", format="latex")
   print(re.search(r"<m:rad>.*?</m:rad>", d.part_xml("word/document.xml")).group(0))
   # observed: <m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:e>...</m:e></m:rad>
   # expected: <m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>...</m:e></m:rad>
   ```
2. **`add_footnote` writes the note's number run with `w:rPr` after `w:footnoteRef`** (schema: `w:rPr` first), and
   no separator / continuation-separator notes; the note text has no space after the number ("1Source").
   ```python
   import rdocx
   d = rdocx.Document(); p = d.add_paragraph("x"); d.add_footnote(p, "note")
   print(d.part_xml("word/footnotes.xml")[-200:])
   # observed: <w:r><w:footnoteRef/><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr></w:r>
   # expected: <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
   ```
   A reference placed after a run in the middle of a paragraph works (`add_footnote(run)` after `docx_ops.isolate`).
3. **Hyperlinks never get the Hyperlink style**: `add_hyperlink`, `add_heading_hyperlink`, Markdown import
   (`rdocx convert x.md`, `from_markdown`, also with a template that defines `Hyperlink`) and `from_html` write
   `<w:hyperlink><w:r><w:t>` with no `w:rStyle`, colour or underline: the link reads as plain text in Word and
   LibreOffice. Workaround: colour and underline the returned run (D10), or the runs after an import (D13, 7 lines).
   ```python
   import rdocx
   t = rdocx.Document(); t.add_style("Hyperlink", style_type="character", color="0563C1", underline=True)
   d, _ = rdocx.Document.from_markdown("See [site](https://example.com).", template=t)
   print("rStyle" in d.part_xml("word/document.xml"))   # False; expected the Hyperlink style on the link run
   ```
4. **`add_heading_hyperlink` retires the paragraph handle** it appends to (`StaleElementError ... First invalidated by
   Document.add_heading_hyperlink`), whereas `references/python-api.md` (Handles) lists it among the appends that keep
   paragraphs valid. Re-fetching `doc.paragraphs[i]` works. Doc or behaviour to align.
   ```python
   import rdocx
   d = rdocx.Document(); h = d.add_heading("Target", 1); p = d.add_paragraph("See ")
   d.add_heading_hyperlink(p, d.paragraphs[0], "Target")
   p.add_run(".")   # raises StaleElementError
   ```
5. **`Cell.merge` keeps the empty paragraph of an empty cell** it absorbs, so the merged row grows by a blank line;
   Word and python-docx drop it. `set_cell_grid_span` gives one paragraph.
   ```python
   import re, rdocx
   d = rdocx.Document(); t = d.add_table(1, 2); t.cell(0, 0).text = "Total"
   d.tables[0].cell(0, 0).merge(d.tables[0].cell(0, 1))
   print(len(re.findall(r"<w:p[ />]", d.tables[0].cell(0, 0).xml)))   # 2; Word: 1
   ```
6. **`Table.insert_column` creates bare cells** (a `w:tcW` and an empty `w:p`): no paragraph or run properties are
   copied from the neighbouring column, unlike Word's Insert Column; the header's bold is set by hand (one line).
7. **Each structural edit costs in proportion to the styles part, not only the body**: `clone_row` takes 28 ms on a
   one-table document made from python-docx's template (349 KB `styles.xml`), 5 ms with a minimal styles part, 1 to
   2 ms on `rdocx.Document()`. D04 (5 `clone_row`, 18 `cell.replace_text`) takes 1.1 s against 0.07 s for the
   other stacks. Reproduction: build `python-docx Document()` + `add_table(3, 3)`, then time
   `doc.tables[0].clone_row(1)` with rdocx, before and after cutting `styles.xml` down to the styles used.
   `references/gaps.md` describes costs that grow with the body only.
8. `rdocx validate` passes outputs that break the schema (findings 1 and 2); `references/cli.md` calls it a check of
   "package and schema invariants", which agents may read as an XSD check.
