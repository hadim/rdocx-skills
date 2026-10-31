# Validity: which checks catch a broken .docx/.pptx, which generators refuse to write one

Results: `results/validity.json` (regenerate with `.cache/venv/bin/python validity/run.py`, about 40 s on
32 CPUs). Versions are in the JSON (`versions`): rdocx 0.16.4 / rpptx 0.14.4 (pin 401e3f42), python-docx
1.2.0, python-pptx 1.0.2, docx-js 9.15.0, pptxgenjs 4.0.1, pandoc 3.13, LibreOffice 26.2.6.

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

### Bases (Part 0)

| Base | Spec oracle | Flagged by |
|---|---|---|
| docx, rdocx | invalid: `w14:paraId` on comment paragraphs without `mc:Ignorable="w14"` (Part 3) | none |
| docx, docx-js | valid | none |
| docx, python-docx | invalid: `w:zoom` without the required `w:percent` (python-docx's default template) | validate.py (zoom); rdocx validate (for another reason: `CommentReference`/`CommentText` styles that python-docx 1.2's `add_comment` uses but does not define) |
| pptx, rpptx | valid | none |
| pptx, pptxgenjs | invalid: `notesMasterIdLst` after `sldIdLst`; three `c:axId` in `c:barChart` (max 2) | rpptx open (`read chart data failed: ChartML contains duplicate c:axId`) |
| pptx, python-pptx | invalid: negative `c:axId` values (xs:unsignedInt) | none |

No validator flags the four pptx/docx schema defects of the pptxgenjs and python-pptx charts or of the rdocx
comments: validate.py does not schema-check pptx chart parts at all and reorders `p:presentation` before
validating; rpptx validate does not check either.

### Detection (Part 1), (mutation, base) pairs on bases the checker passes

docx (31 invalid mutations, 12 controls):

| Checker | All invalid | zip | package | xml | schema | reference | semantic | FP on controls |
|---|---|---|---|---|---|---|---|---|
| rdocx validate | 26/62 (42 %) | 4/6 | 14/14 | 4/8 | **0/20** | 4/12 | 0/2 | 2/24 |
| rdocx open | 30/93 (32 %) | 6/9 | 12/21 | 6/12 | 0/30 | 6/18 | 0/3 | 0/36 |
| validate.py | 52/62 (84 %) | 4/6 | 12/14 | 8/8 | 18/20 | 8/12 | 2/2 | 0/24 |
| validate.py --original | 78/93 (84 %) | 6/9 | 18/21 | 12/12 | 27/30 | 12/18 | 3/3 | 0/36 |
| python-docx open | 33/93 (36 %) | 6/9 | 15/21 | 9/12 | 0/30 | 3/18 | 0/3 | 3/36 |
| LibreOffice | 32/93 (34 %) | 9/9 | 9/21 | 9/12 | 0/30 | 5/18 | 0/3 | 0/36 |
| pandoc | 12/93 (13 %) | 6/9 | 3/21 | 3/12 | 0/30 | 0/18 | 0/3 | 0/36 |
| stack rdocx (validate + open) | 28/62 (45 %) | 4/6 | 14/14 | 4/8 | 0/20 | 6/12 | 0/2 | 2/24 |
| stack builtin (validate.py + LO + pandoc) | 55/62 (89 %) | 6/6 | 12/14 | 8/8 | 18/20 | 9/12 | 2/2 | 0/24 |

pptx (32 invalid mutations, 11 controls, 1 office-only):

| Checker | All invalid | zip | package | xml | schema | reference | semantic | FP on controls |
|---|---|---|---|---|---|---|---|---|
| rpptx validate | 63/94 (67 %) | 6/9 | 15/21 | 6/12 | 21/31 | 15/18 | 0/3 | 0/33 |
| rpptx open | 34/63 (54 %) | 4/6 | 6/14 | 4/8 | 12/21 | 6/12 | 2/2 | 0/22 |
| validate.py | 66/94 (70 %) | 6/9 | 18/21 | 12/12 | 15/31 | 15/18 | 0/3 | 0/33 |
| python-pptx open | 37/94 (39 %) | 6/9 | 13/21 | 9/12 | 0/31 | 9/18 | 0/3 | 3/33 |
| LibreOffice | 15/94 (16 %) | 6/9 | 6/21 | 3/12 | 0/31 | 0/18 | 0/3 | 0/33 |
| stack rdocx (validate + open) | 46/63 (73 %) | 4/6 | 10/14 | 4/8 | 14/21 | 12/12 | 2/2 | 0/22 |
| stack builtin (validate.py + LO) | 69/94 (73 %) | 9/9 | 18/21 | 12/12 | 15/31 | 15/18 | 0/3 | 0/33 |

Caught by no checker: docx `schema_tc_no_p` (a cell without a paragraph) and `ref_dup_revision_id`; pptx
`opc_bad_target_mode` (TargetMode="Sideways" in presentation.xml.rels), `schema_notes_master_after_slides`,
`schema_unknown_element` (`<p:bogus/>` in a slide's shape tree) and `schema_cnvpr_no_name`.

Readings:

- **docx: validate.py catches 18 of 20 schema-layer pairs, rdocx validate none.** rdocx validate checks the
  package (content types, relationships: 14/14), part well-formedness of the parts the main document relates to,
  style ids and comment ownership, and no WordprocessingML schema rule.
- **pptx: the two validators are complementary.** rpptx validate catches the slide-level value faults
  validate.py lets through (negative shadow `dist`, `algn="middle"`, negative `a:ext`), because validate.py
  reports a slide XSD error only when it is on its 7-item denylist; validate.py catches what rpptx validate
  misses (presentation.xml reordered, a slide without its content-type Override, undeclared chart axes,
  well-formedness of docProps/app.xml).
- **Opening is not checking.** python-docx and python-pptx caught 0 schema faults; LibreOffice converted
  every schema-invalid file to PDF (it repairs silently), refused 9/9 broken ZIPs in docx but rendered a pptx
  whose first slide was truncated.
- **No false positive on the controls** for any validator but rdocx validate's undefined-style judgement
  (2/24). MCE AlternateContent, ignorable unknown namespaces, extension lists, custom XML parts, stored ZIP
  items, `[Content_Types].xml` last, a 200,000-character paragraph, empty cell paragraphs, the largest legal
  slide id and .dotx/.potx content types passed everywhere (python libraries aside on templates).

### Prevention (Part 2)

Outcomes per generator (`prevention_summary`): refused at the call / written invalid / written valid.

| Generator | Refused at call | Written invalid | Written valid (same as the correct call) | Written valid (different) |
|---|---|---|---|---|
| docx-js (12 cases) | 1 | 3 | 2 | 6 |
| pptxgenjs (18) | 0 | 10 | 2 | 6 |
| rdocx (15) | 13 | 0 | 2 | 0 |
| rpptx (15) | 10 | 1 | 1 | 3 |
| python-docx (10) | 5 | 2 | 0 | 3 |
| python-pptx (8) | 2 | 3 | 0 | 3 |

- **pptxgenjs writes 10 invalid files and refuses none**: '#' or 8-digit colours on shadows and chart
  gridlines, a negative shadow offset, a shared shadow object (second shape gets `dir`/`alpha` converted twice:
  `dir="162000000000"`, out of range), a scheme colour in a hex-only option (`srgbClr val="bg2"`), '&' in
  `pres.company` or a layout title (not well-formed), secondary axes without `valAxes`/`catAxes` (undeclared
  axis ids), U+0001 in text. Caught by: validate.py 7 of 10, rpptx validate 6, rpptx open 6, python-pptx 2,
  LibreOffice 1; **the '#' gridline colour in a chart is caught by nothing** (no checker schema-validates chart
  parts; rpptx open fails on that deck, but equally on the correct one, for pptxgenjs' third `c:axId`). On fills
  pptxgenjs fixes '#' and 8 digits silently (the 8-digit fill loses its alpha; text, line, table and background
  colours behaved the same in a side probe, not recorded in the JSON). `stacked outEnd` is spec-valid
  and flagged by validate.py only; letterSpacing, rectRadius on RECTANGLE, a misspelt placeholder and an
  off-slide box are silently ignored or written as asked.
- **docx-js** refuses an 8-digit colour, strips '#', drops U+0001 silently, and writes three invalid files: a
  `PageBreak` outside a paragraph (a `w:r` in `w:body`), a comment reference to a missing comment, a negative
  picture width; validate.py catches all three, rdocx validate the comment one.
- **rdocx refuses 13 of 15** with an error naming the fix: 8-digit colour, negative picture size, bare int font
  size, malformed / misordered / unknown-element / bad-enum XML through `replace_xml`, unknown style, stale run
  handle (`StaleElementError`), comment on missing text, wrong replacement count, U+0001 in text, CLI output
  over its input. It accepts '#' colours and normalises them.
- **rpptx refuses 10 of 15** the same way (8-digit colour, negative shadow distance, bare int size, the four
  `replace_xml` faults, stale paragraph handle, wrong count, CLI output over input), escapes '&' in the company
  and a layout name and writes U+0001 as `_x0001_`, but **accepts a negative text box width and writes it**
  (`a:ext cx="-2743200"`; rpptx validate flags the saved file).
- **python-docx / python-pptx** refuse '#' only by accident (`int('#F', 16)`), turn 8 digits into a wrong
  colour (`FF000080` -> `FF0080`, `00000020` -> `000020`), python-docx writes `w:sz="0"` for `font.size = 12`
  (python-pptx refuses its `sz=0`), python-pptx accepts a negative width, and both lose an edit made through a
  stale run handle without a word. Raw lxml writes whatever it is given (pPr after r, `w:bogus`, `p:bogus`):
  no checker tells those python-docx files from the correct call (validate.py flags both, for the template's
  `w:zoom`; rdocx validate passes both), and nothing flags the python-pptx `p:bogus`.

Atomic save (`atomic_save`; the process may write at most 4 KiB, so the save fails midway):

| Save | Input after the failure |
|---|---|
| rdocx `Document.save(input)`, rpptx `Presentation.save(input)` | intact (both write a temporary file and rename it) |
| `docx_ops.save_atomic` / `pptx_ops.save_atomic` over an existing output | intact, no temporary file left |
| python-docx `Document.save(input)` | **replaced by a valid-looking 4 KiB ZIP that lacks styles, settings, theme, comments and the picture** (zipfile writes the central directory of what it had written) |
| python-pptx `Presentation.save(input)` | truncated to 4 KiB, not a ZIP |
| docx-js `writeFileSync` / pptxgenjs `writeFile` over an existing output | truncated to 4 KiB, not a ZIP (the input is never written) |
| built-in edit flow `rm -f ../out.docx && zip -Xr ../out.docx .` | input intact (never written); the previous output is gone, nothing replaces it |

## Caveats

- Mutations are single-fault and synthetic; real corruption often stacks faults. Each mutation runs on three
  bases, which is why pair counts are multiples of 2 or 3.
- Detection says nothing about what Word or PowerPoint do; only the cited SKILL.md lines speak to that.
- The oracle covers zip, OPC, XML and XSD rules, not the prose rules; it does not check core.xml against its
  schema (the Dublin Core imports are not installed) nor Microsoft extension parts.
- LibreOffice is judged on "a PDF came out", the only signal its batch mode gives.
- Parallel timing (16 jobs) inflates every `seconds` value somewhat.

## rdocx findings

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
