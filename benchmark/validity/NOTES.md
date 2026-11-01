# Validity: which checks catch a broken .docx/.pptx, which generators refuse to write one

Results: `results/validity.json` (regenerate with `.cache/venv/bin/python validity/run.py`, about 40 s on
32 CPUs). Versions are in the JSON (`versions`).

## Method

Files of this section:

| File | Role |
|---|---|
| `gen/` | one base document and one base deck per generator (rdocx/rpptx, docx-js/pptxgenjs, python-docx/python-pptx), each written the way its skill shows: styles, a heading, a table, a picture, a comment and a tracked change (docx); two layouts, a table, a native chart, a picture, speaker notes (pptx) |
| `mutations.py` | 43 docx and 44 pptx mutations on the raw package (zipfile + lxml, or bytes): 31 + 32 `invalid`, 12 + 11 negative controls (`valid`), 1 pptx spec-valid file PowerPoint refuses (`office_only`) |
| `oracle.py` | the spec oracle (zipfile + lxml only): ZIP integrity, OPC (content types, relationships, target content types, rels and content-types XSD), XML 1.0 well-formedness, and the ISO/IEC 29500-4:2016 transitional XSDs after Markup Compatibility preprocessing |
| `checkers.py`, `probe_*.py` | the checks an agent runs, each as its stack runs it |
| `prevention.py` | Part 2: footguns and their analogues, refused or written; atomic save |
| `run.py` | runs everything, writes the JSON |

**Checkers.** `rdocx_validate` (`rdocx validate F` / `rpptx validate F`, flagged on a non-zero exit),
`rdocx_open` (the Python binding loads the file and reads text, tables, comments, chart categories, notes),
`builtin_validate` (`scripts/office/validate.py F` from the installed skill, run with `common.PY`),
`builtin_validate_original` (the same with `--original BASE`, the skill's edit flow; an extra column),
`pylib_open` (python-docx / python-pptx load and iterate), `libreoffice` (PDF conversion through the skill's
`soffice.py`; flagged only when no PDF comes out), `pandoc` (`pandoc -f docx -t markdown`, docx only). The
`stack:*` rows are the union of what each stack's workflow runs: rdocx = validate + open; builtin = validate.py
+ LibreOffice (+ pandoc for docx); pylib = open.

**Counting.** Every mutation is applied to the three bases of its format. A (mutation, base) pair counts for a
checker only when that checker passes the unmutated base (otherwise it would flag the mutation for the base's
own reason); `bases_used` in the summary says which bases each checker kept. Detection = flagged pairs / pairs
on `invalid` mutations; false positives = flagged pairs on `valid` controls. `office_only` is reported apart.
`seconds` is wall clock of one command, process start included, measured while 16 checks ran in parallel
(indicative only; medians: rdocx validate 6 to 12 ms, rdocx open 30 ms, validate.py 115 to 170 ms, python
open 60 to 90 ms, LibreOffice 0.5 to 0.6 s, pandoc 11 ms).

## Label sources

- `expected` is set by construction from ECMA-376: Part 1 (the transitional schemas of ISO/IEC 29500-4:2016 and
  the prose "shall" rules), Part 2 (OPC), Part 3 (MCE), plus XML 1.0 and the ZIP format. Each record has
  `rule` (the clause and why it breaks) and `basis`: `zip`, `opc`, `xml`, `schema`, `mce`, `prose` or
  `judgement`. Clause numbers are cited from memory of ECMA-376 Part 1 (5th edition) and were not re-checked
  against the text in this run; the wording of each rule is what the label rests on.
- Every label with basis zip / opc / xml / schema / mce was confirmed on every base by the oracle
  (`label_confirmed_by_oracle`: new oracle issues for an invalid mutation, none for a control). Prose and
  judgement labels are not machine-checked.
- `office` holds what Word or PowerPoint does only when a source says so. The only sources found are the
  built-in pptx SKILL.md (8-digit and '#' colours, negative shadow offset, `outEnd` on stacked bars, undeclared
  secondary axes, reordering `p:presentation`, `notesMasterIdLst` after `sldIdLst` "read happily") and the
  built-in validator's own list of "slide-XML schema errors PowerPoint refuses the file over"
  (`helpers/pptx_slide.py`). For '&' in app.xml or a layout name the SKILL.md only says the part is
  unparseable and validate.py fails; it says nothing of PowerPoint, and the record says so. Everything else is
  `unknown`; no docx mutation has a source.
- The oracle fetches its XSD files from their publishers (`validity/schemas.py`, pinned by SHA-256): the
  transitional schemas of ECMA-376 Part 4, 5th edition (the text of ISO/IEC 29500-4:2016), the OPC schemas of
  Part 2 and W3C's `xml.xsd`. It strips every `mc` attribute and resolves `AlternateContent` before validating.

## Judgement calls

1. **Transitional schemas.** Run properties are an `xsd:choice` in ISO/IEC 29500-4:2016 (`EG_RPrBase`), so
   reordering rPr children is valid there (the 1st edition had a sequence): it is a negative control
   (`docx.ctl_rpr_reordered`). The paragraph-mark rPr keeps its `w:ins`/`w:del` first (`CT_ParaRPr`), so
   `docx.schema_pmark_del_misordered` is invalid.
2. **A table cell with no paragraph** breaks the prose of 17.4.66 ("a table cell shall contain at least one
   block-level element"); the XSD cannot enforce it (an empty `EG_BlockLevelElts` group satisfies `minOccurs=1`),
   so its basis is `prose`, not `schema`.
3. **Undefined style id** (`docx.ctl_style_undefined`): valid by judgement (ST_String; no schema or OPC rule;
   consumers fall back to the default style). rdocx validate calls it an error: that is its only false
   positive, and the reader who counts an undefined style as invalid should move it to the hits.
4. **Dangling internal relationship** (media part deleted, relationship kept): invalid by judgement
   (`opc_media_missing`). **Undeclared chart axis id** and **a table row with fewer cells than grid columns**:
   invalid by judgement (dangling reference; a column without a cell). The oracle checks the first, not the
   others.
5. **`outEnd` labels on a stacked bar chart** are spec-valid (ST_DLblPos has `outEnd`, no clause restricts it),
   so `pptx.office_stacked_outend` is neither a hit nor a false positive: only validate.py flags it (3/3), per
   PowerPoint's behaviour as its SKILL.md reports it.
6. **`notesMasterIdLst` after `sldIdLst`** breaks `CT_Presentation` (invalid), yet the SKILL.md says PowerPoint
   reads it; pptxgenjs writes it that way, so the mutation does not apply to the pptxgenjs base.
7. **Bases are not all spec-valid** (next section). They are used as generated: a checker that flags a base is
   simply not scored on that base.
8. **validate.py without `--original`** is the primary column, as the task specifies (files built from
   scratch); the `--original BASE` column shows the edit flow. With `--original` the python-docx base becomes
   usable (its template error is baselined), so the docx rate is computed on 93 pairs instead of 62, at the
   same 84 %.
9. **python-docx / python-pptx refuse .dotx / .potx** by design (content type check): counted as their 3 + 3
   false positives on the template controls.
10. **pptxgenjs colour footguns were tested on several option paths**: pptxgenjs 4.0.1 strips '#' and replaces
    an 8-digit value with `000000` (console warning) on fills, text and lines, but writes the value raw for
    shadow colours and chart gridlines. Both paths are recorded, so the SKILL.md warning is neither confirmed
    nor refuted wholesale.
11. **docx-js `ImageRun` without `type:`** wrote the same file as with `type: "png"` at 9.15.0 (the SKILL.md
    says it is required): recorded as such.
12. Footguns not exercised (layout or style advice with no file-validity side: A4 default, landscape, TOC heading
    styles, a table as a rule, dot leaders, `isTextBox`, gradients, transparency vs opacity, icons, sections,
    `•` in pptx). Prevention `output_invalid` = new oracle issues compared with the same script done right
    (`control`), plus the judgement checks for undeclared axes and a dangling comment reference.

## Results

The tables of `benchmark/README.md`, generated from `results/validity.json` by `report.py`. This file does not repeat the numbers: they change with each run.

## Caveats

- Mutations are single-fault and synthetic; real corruption often stacks faults. Each mutation runs on three
  bases, which is why pair counts are multiples of 2 or 3.
- Detection says nothing about what Word or PowerPoint do; only the cited SKILL.md lines speak to that.
- The oracle covers zip, OPC, XML and XSD rules, not the prose rules; it does not check core.xml against its
  schema (the Dublin Core imports are not installed) nor Microsoft extension parts.
- LibreOffice is judged on "a PDF came out", the only signal its batch mode gives.
- Parallel timing (16 jobs) inflates every `seconds` value somewhat.

## rdocx findings

Met on the first run (2026-10-31, pin `401e3f42`). They were reported upstream as tensorbee/rdocx#431 to #437, and the pin the benchmark now runs on carries the fixes of all but the open gaps of `tests/gaps.py`. Kept as the record of what the benchmark found; the README has the current results.

Each with a reproduction that builds its own input (run with `$R/python`, R the rdocx `bin` folder).

1. **`rdocx validate` checks no WordprocessingML schema rule** (0/20 schema-layer pairs; validate.py 18/20):
   w:pPr after w:r, an unknown `w:bogus`, `w:jc w:val="middle"`, text directly in `w:r`, `w:sectPr` not last,
   `w:color w:val="#FF0000"`, a negative `wp:extent`, `w:ins` without `w:id` all pass with exit 0. Its help
   says "Validate OOXML conformance" and cli.md "package and schema invariants".
   ```python
   import subprocess, sys, zipfile, rdocx
   d = rdocx.Document(); d.add_paragraph("Hello"); d.save("base.docx")
   zin = zipfile.ZipFile("base.docx")
   with zipfile.ZipFile("bad.docx", "w", zipfile.ZIP_DEFLATED) as z:
       for i in zin.infolist():
           x = zin.read(i)
           if i.filename == "word/document.xml":
               x = x.replace(b"<w:p><w:r>", b'<w:p><w:bogus/><w:pPr><w:jc w:val="middle"/></w:pPr><w:r>')
           z.writestr(i, x)
   print(subprocess.run([sys.argv[1], "validate", "bad.docx"]).returncode)   # 0 (expected: non-zero)
   ```
2. **rdocx reads and re-saves U+0001 in document.xml** (not well-formed XML 1.0): `rdocx validate` passes it
   before and after `Document(...).save(...)`, and the saved file still holds the byte. The writing side refuses
   the same character (`add_paragraph("bad\x01char")` raises XmlError), so only a damaged input propagates.
   Reproduction: as above with `x.replace(b"Hello", b"He\x01llo")`, then `rdocx.Document("bad.docx").save("out.docx")`
   and `b"\x01" in zipfile.ZipFile("out.docx").read("word/document.xml")` -> True.
3. **`rdocx validate` / `rpptx validate` do not parse docProps/app.xml**: `<Company>Harrow & Sons</Company>`
   passes both (validate.py fails it; LibreOffice refuses that pptx outright).
4. **Duplicate ZIP entries pass both validators, and readers disagree on which copy is the document**: with two
   `word/document.xml` items, rdocx and python-docx read the last, pandoc the first, LibreOffice refuses the
   file. A review made with one tool and an edit with another see different text.
5. **`rdocx validate` misses dangling and duplicate ids**: `a:blip r:embed="rId999"`, a `w:headerReference`
   whose `r:id` names no relationship (the binding itself fails to open that file), a `w:footnoteReference` to
   a missing footnote, two revisions with one `w:id`, and `w:t` inside `w:del` (should be `w:delText`).
6. **`rpptx validate` misses** an unknown `p:` element in a slide, `p:cNvPr` without `name`, `p:sldSz` moved
   before `p:sldIdLst` (the SKILL.md of the built-in pptx skill says that reorder makes a deck unopenable), a
   slide whose Override is removed (content type application/xml), an undeclared chart axis id, and a table row
   with fewer cells than `a:gridCol`.
   ```python
   import subprocess, sys, zipfile, rpptx
   prs = rpptx.Presentation(); prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "Hello"; prs.save("b.pptx")
   zin = zipfile.ZipFile("b.pptx")
   with zipfile.ZipFile("bad.pptx", "w", zipfile.ZIP_DEFLATED) as z:
       for i in zin.infolist():
           x = zin.read(i)
           if i.filename == "ppt/slides/slide1.xml":
               x = x.replace(b"</p:spTree>", b"<p:bogus/></p:spTree>")
           z.writestr(i, x)
   print(subprocess.run([sys.argv[1], "validate", "bad.pptx"]).returncode)   # 0 (expected: non-zero)
   ```
7. **rpptx cannot read any pptxgenjs bar chart, yet `rpptx validate` passes the deck**: `shape.chart.categories`
   raises `RpptxError: read chart data failed: ChartML contains duplicate c:axId` on a one-series pptxgenjs bar
   chart. The ids are not duplicates: pptxgenjs writes three distinct `c:axId` in `c:barChart` (the schema allows
   two). The message misleads, and validate and the reader disagree.
   ```js
   // NODE_PATH=... node f.js ; then $R/rpptx validate chart.pptx (exit 0) and read shape.chart.categories (raises)
   const P = require("pptxgenjs"); const p = new P();
   p.addSlide().addChart(p.charts.BAR, [{ name: "A", labels: ["x", "y"], values: [1, 2] }], { x: 1, y: 1, w: 6, h: 4 });
   p.writeFile({ fileName: "chart.pptx" });
   ```
8. **rpptx `add_textbox` accepts a negative width** and writes `a:ext cx="-2743200"`, while rdocx `add_picture`
   and rpptx `shadow.distance` refuse negatives: `slide.shapes.add_textbox(Inches(1), Inches(1), -Inches(3),
   Inches(1))` then `save`; `rpptx validate` flags the saved file.
9. **rdocx writes `w14:paraId` on comment paragraphs without `mc:Ignorable="w14"`** on `w:comments` (and no
   `mc` declaration): a consumer that does not understand w14 must reject the attribute (Part 3). Low impact,
   since Word understands w14; found by the oracle on the rdocx base (`docx_ops.comment_on_text`).
10. **Skills documentation**: `pptx_ops.save_atomic`'s docstring says "rpptx's own save() writes in place". At
    this pin both `Document.save` and `Presentation.save` write a temporary file and rename it (new inode; a
    save into a read-only folder fails; after a failed write the input is intact). The SKILL.md advice to use
    `save_atomic` still holds for its refusal of `out == src`.
11. **Judgement, not a spec violation**: `rdocx validate` reports an undefined style id as an error, so python-docx
    1.2's own `add_comment` output (`CommentReference`/`CommentText` not defined) fails it.

What rdocx and rpptx do well here, for balance: they refuse most mistakes at the call with an error that names
the fix (23 of 30 cases, against 7 of 18 for python-docx/python-pptx and 1 of 30 for docx-js/pptxgenjs),
`replace_xml` refuses malformed, misordered, unknown and out-of-enum XML, stale handles raise instead of losing
edits, saves are atomic, and rpptx validate is the only validator here that catches slide-level value faults
(negative shadow offset, bad alignment enum, negative extent).

## Built-in stack notes (not rdocx, for the report)

- validate.py reports a slide XSD error only when it matches one of 7 denylisted messages, and skips pptx chart
  parts: a negative shadow offset (which its own SKILL.md says corrupts the file) and `algn="middle"` pass it.
- validate.py fails every python-docx document (default template's `w:zoom` without `w:percent`) unless
  `--original` is given, so on files built from scratch with python-docx it cannot tell a new fault from the
  template's.
