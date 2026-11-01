# Word review, collaboration, layout and rendering tasks (R01-R15)

Can each stack do the task, with how much code, how many tools and processes, and how much hand-written OOXML?
Three stacks, as `README.md` defines them: `rdocx` (the pinned build, CLI, Python and
`docx_ops.py`), `builtin` (Claude's built-in docx skill: unzip / edit XML / zip with its helper scripts, pandoc to
read, LibreOffice through `soffice.py` and pdftoppm to render), `pylib` (python-docx 1.2.0, lxml where it has no
API). The authoring tasks D01-D14 are another section's (`NOTES.md` of this folder, if any).

## Running

```bash
.cache/venv/bin/python features/docx/run_review.py            # all tasks, timings included: about 50 s here
.cache/venv/bin/python features/docx/run_review.py --tasks R08 --repeat 0 --no-save   # one task, no timing
```

It writes `results/features_docx_review.json`. Scratch files go to `.cache/work/features_docx_review/<task>/`
(`input/`, then one folder per stack holding its output and LibreOffice's PDF of it).

## Method

- **Inputs** (`Rxx-*/make_input.py DIR`, helpers in `review_build.py`): written part by part as Word 2016 writes
  a package (styles with docDefaults and the built-in styles, settings with rsids, theme, font table, app and core
  properties, `w14:paraId` on paragraphs, text cut into several runs with rsids and `proofErr` marks). Comments come
  as Word writes them (`comments.xml`, `commentsExtended.xml` with `w15:done`, `commentsIds.xml`, `people.xml`, a
  reply's markers nested in its parent's), tracked changes as `w:ins` / `w:del` / `w:rPrChange` with `w:id`,
  `w:author`, `w:date`, a deleted or inserted paragraph with its tracked paragraph mark. Fixed zip timestamps: the
  same script writes the same bytes. Every input passes `rdocx validate`, the skill's `validate.py` and opens in
  LibreOffice. The same input goes to the three stacks.
- **Implementations**: `rdocx.sh|py`, `builtin.sh|py` (with a `builtin_*.py` for the XML edit when the skill's
  workflow needs one), `pylib.py`; each takes the input path(s) then the output path. A `# bench: {...}` header
  declares the tools, the processes one run starts, `raw_xml`, the other files it needs and, where the stack
  cannot do the task, `unsupported` with the reason. No implementation runs a validator itself (the runner runs all
  of them on every output), so that no stack pays for its verification step and another not.
- **Oracle** (`Rxx-*/check.py OUTPUT`, helpers in `review_oracle.py`): zipfile, lxml and the standard library only.
  Text views follow the OOXML semantics: accepted (insertions in, deletions out, a deleted paragraph mark joins
  its paragraph to the next) and rejected (the reverse). Effective formatting walks docDefaults, the paragraph and
  character style chains and direct formatting. PNG pages are decoded with zlib (size, share of non-white pixels,
  rows that differ); PDF pages are counted from the page objects, object streams included, and checked for text
  or drawing operators. Exit 0 and `"ok": true` is a pass.
- **Validity** (docx outputs): `rdocx validate OUT`; the skill's `scripts/office/validate.py OUT --original IN`
  (`--author "Dana Reviewer"` for the redlines R04 and R08, as the skill says); `libreoffice_opens`: `soffice.py`
  converts the output to a non-empty PDF.
- **LOC**: non-blank lines that are not comments or docstrings, over every file the implementation needs (not
  `make_input.py`, not `check.py`, not the skill's own scripts it calls). An `unsupported` stack counts 0, except
  R09 and R10 for pylib, whose fallback is run and counted (below).
- **Processes**: the programs one run starts that do work (interpreters, CLIs, LibreOffice, pdftoppm, pdftotext,
  pandoc, unzip, zip), a helper's own LibreOffice included; the shell and coreutils (`mktemp`, `cp`, `sed`) not.
- **Seconds**: wall time of the whole implementation as an agent runs it, median of 5 runs after a warm-up (3 for
  the ones that start LibreOffice), measured one task at a time after the parallel correctness pass.

## Results

The tables of `benchmark/README.md`, generated from `results/features_docx_review.json` by `report.py`. This file does not repeat the numbers: they change with each run.

## Judgement calls

- **R01, R02**: the built-in `comment.py` stamps the current time and has no `--date`, so the script rewrites the
  timestamp it wrote; it prints the marker snippet, which the script places (splitting the run that
  `merge_runs.py` made), as the SKILL.md says. Resolving has no helper: `w15:done="1"` is set by hand.
- **R02 oracle**: besides the thread link (`w15:paraIdParent`) and `w15:done`, the reply must be referenced from
  the document (a `w:commentReference`, its range next to the parent's). Word writes one for every reply, and
  LibreOffice drops a comment that the document never references (checked: LibreOffice's ODT of the rdocx output
  holds two annotations, of the built-in and python-docx outputs three, the reply threaded). This is the check
  rdocx fails; how Word itself shows a reply without a reference was not checked (no Word here).
- **R03, R05**: the built-in stack reads with pandoc, the reader its SKILL.md names (`--track-changes=all -t
  json`). For R05 that misses formatting changes: pandoc reports insertions and deletions only. A raw-XML listing
  (what the pylib version does) would pass; it is not what the skill shows first.
- **R04**: rdocx has no tracked edit call; its documented recipe ("Redline two versions") is `rdocx replace` into
  a copy, then `rdocx compare --granularity word`.
- **R06**: the built-in skill documents that `accept_changes.py` leaves an empty paragraph when a deleted
  paragraph is followed by an empty spacer paragraph. The input has that case twice (a plain paragraph, and the
  last item of a bulleted list); LibreOffice 26.2 joined both correctly, so the built-in stack passes.
- **R07, R08, R15**: no built-in helper rejects by author, redlines two files or merges documents, so the built-in
  versions are Python + lxml on the unzipped XML (difflib for the redline: paragraphs aligned, then a word diff
  inside each changed paragraph). The python-docx versions are the same logic over python-docx's element tree.
  These are written with the inputs' traps in view (the docDefaults of R15 in particular); a first attempt by an
  agent could well miss them, as rdocx's own call does.
- **R08 oracle**: accepted view == v2 and rejected view == v1 are necessary but not enough for "word level": a word
  that a word diff leaves unchanged, and that is not next to a changed word, must not sit in a revision. Folding a
  short unchanged word into a changed stretch (rdocx's documented behaviour) is allowed by the neighbour rule.
- **R09, R10**: for the built-in stack, page numbers come from LibreOffice: `soffice.py` to PDF, `pdftotext -layout`,
  then each heading (or the text before each PAGE field) looked up page by page, and the TOC entries or field
  results written into the XML by hand. Another path an agent may know: a LibreOffice round trip
  (`soffice.py --convert-to docx`) refreshes the fields of R10 (it passes this oracle and `validate.py`) but
  rewrites the whole package through LibreOffice's model; not taken, as the skill does not show it.
  python-docx cannot paginate: its usual fallback, run and recorded as `unsupported` with its LOC, marks the TOC
  dirty / writes today's date and sets `w:updateFields`, so that Word refreshes everything on open after a prompt.
- **R11, R13**: LibreOffice's PDF shows tracked changes by default (coloured, underlined, struck through), so the
  built-in tracked view is a plain conversion and the accepted view needs `accept_changes.py` first. rdocx renders
  the accepted view by default and the tracked one with `--revision-view tracked`. PNG at 100 dpi with
  `pdftoppm -png -r 100` (the SKILL.md shows `-jpeg`; the task asks for PNG). rdocx gets `--font-dir` with the
  Carlito and Caladea fonts, LibreOffice gets them through `FONTCONFIG_FILE`.
- **R12**: rdocx's `doc.layout()` gives the page of every body block; the built-in stack takes the headings from
  pandoc's AST and finds each one as a line of its own in LibreOffice's PDF text. python-docx's
  `rendered_page_breaks` only reads breaks Word cached at its last save, and these files were not saved by Word.
- **R14**: rdocx `diff --json` reshaped; pandoc's plain text and python-docx's paragraphs through difflib. pandoc
  puts a table row on one line, so the oracle accepts a reported entry that contains the changed paragraph.
- **R15**: "keep B's formatting" includes B's document defaults (Georgia body text in `w:docDefaults`): A and B use
  the same style ids with different definitions, the same `numId` for lists of different formats, and the same
  picture part name with different pixels.

## Caveats

- Ground truth for pages (R09, R10, R12) is by design: headings start a page after a page break, or fall at least
  a paragraph and a half from a page edge; rdocx and LibreOffice agree on every one. Word itself was not run.
- The PNG and PDF oracles check page count, size, ink and that the two views differ; they cannot read the text, so
  they do not prove that the accepted view shows the accepted text (the PDFs were spot-checked with pdftotext: both
  stacks' `document.pdf` read "covered four sites", the accepted wording).
- `accept_changes.py` uses one fixed LibreOffice profile (`/tmp/libreoffice_docx_profile`): two runs at once, from
  this runner or another section's, would collide. The runner serialises its own; it cannot see other sections'.
- The R02 interop check uses LibreOffice as the witness for a missing reply reference; Word was not available.
- Timings are sub-second for every non-LibreOffice implementation and mostly measure interpreter starts.

## rdocx findings

Met on the first run (2026-10-31, pin `401e3f42`). They were reported upstream as tensorbee/rdocx#431 to #437, and the pin the benchmark now runs on carries the fixes of all but the open gaps of `tests/gaps.py`. Kept as the record of what the benchmark found; the README has the current results.

Each reproduction builds its own input with python-docx 1.2 and runs the pinned build
(`~/.local/share/rdocx-skills/current/bin`, pin `401e3f42`, rdocx 0.16.4).

### 1. A reply gets no range or reference in the document (R02)

`Document.reply_to` and `rdocx comment reply` write the reply into `comments.xml` and link it in
`commentsExtended.xml` (`w15:paraIdParent`), but put no `w:commentRangeStart` / `w:commentRangeEnd` /
`w:commentReference` for it in `document.xml`, add no `commentsIds.xml` entry, and take no initials. Word writes a
range and a reference for every reply; LibreOffice drops the reply on import (its ODT holds the root comment only).
`rdocx validate` and the skill's `validate.py` both pass the file.

```python
import os, subprocess, tempfile, zipfile
import docx
d = tempfile.mkdtemp()
doc = docx.Document()
p = doc.add_paragraph("Pay within thirty days.")
doc.add_comment(p.runs, text="Why thirty?", author="A", initials="A")
src, out = os.path.join(d, "in.docx"), os.path.join(d, "out.docx")
doc.save(src)
R = os.path.expanduser("~/.local/share/rdocx-skills/current/bin/rdocx")
subprocess.run([R, "comment", "reply", src, "--id", "0", "--author", "B", "--text", "Agreed",
                "--date", "2026-10-30T10:00:00Z", "-o", out], check=True, capture_output=True)
z = zipfile.ZipFile(out)
print("comments:", z.read("word/comments.xml").decode().count("<w:comment "),
      "references:", z.read("word/document.xml").decode().count("<w:commentReference"), "(expected 2)")
# prints: comments: 2 references: 1 (expected 2)
```

Expected: the reply's range and reference next to its parent's, as Word writes them (the built-in `comment.py`
prints that layout).

### 2. A word-level compare marks unchanged words when the edited file splits them into other runs (R08)

`rdocx compare --granularity word` (and `Document.compare(..., granularity="word")`) shows a word as deleted and
inserted again when the edited file holds it split across runs, even when the two texts are identical. In R08 an
unchanged paragraph of v2 got revisions ("accrues"), as did "invoice" and "letter". `--granularity run` is worse,
`--granularity character` passes the R08 oracle.

```python
import os, subprocess, tempfile
import docx
d = tempfile.mkdtemp()
a = docx.Document(); a.add_paragraph("Each invoice is paid within thirty days.")
a.save(os.path.join(d, "a.docx"))
b = docx.Document(); p = b.add_paragraph("Each invoi"); p.add_run("ce is paid within thirty days.")
b.save(os.path.join(d, "b.docx"))                       # same text, "invoice" split over two runs
R = os.path.expanduser("~/.local/share/rdocx-skills/current/bin/rdocx")
subprocess.run([R, "compare", f"{d}/a.docx", f"{d}/b.docx", "--author", "X", "--timestamp",
                "2026-10-30T12:00:00Z", "--granularity", "word", "-o", f"{d}/r.docx"], check=True, capture_output=True)
print(subprocess.run([R, "revision", "list", f"{d}/r.docx"], capture_output=True, text=True).stdout)
# 3 revisions (one deletion, two insertions) around "invoice"; expected none
```

### 3. `insert_document` drops the inserted document's `w:docDefaults` (R15)

B's body font is set in its `w:docDefaults` (Georgia). After `a.insert_document(b, at=...)` B's paragraphs use A's
defaults (Calibri): conflicting styles are renamed (`Heading1Merge1`), lists renumbered and pictures copied, but
nothing carries B's defaults. `policy=FragmentConflictPolicy.rename_all()` does the same.

```python
import os, subprocess, tempfile, zipfile
import docx
from docx.oxml.ns import qn
d = tempfile.mkdtemp()
a = docx.Document(); a.add_paragraph("A body text."); a.save(f"{d}/a.docx")
b = docx.Document(); b.add_paragraph("B body text.")
fonts = b.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr")).find(qn("w:rFonts"))
for k in list(fonts.attrib):
    del fonts.attrib[k]
fonts.set(qn("w:ascii"), "Georgia"); fonts.set(qn("w:hAnsi"), "Georgia")
b.save(f"{d}/b.docx")
code = (f"import rdocx\na, b = rdocx.Document.open('{d}/a.docx'), rdocx.Document.open('{d}/b.docx')\n"
        f"a.insert_document(b, at=len(a.paragraphs) + len(a.tables))\na.save('{d}/out.docx')")
subprocess.run([os.path.expanduser("~/.local/share/rdocx-skills/current/bin/python"), "-c", code], check=True)
z = zipfile.ZipFile(f"{d}/out.docx")
print("Georgia kept:", any(b"Georgia" in z.read(n) for n in ("word/styles.xml", "word/document.xml")))
# prints: Georgia kept: False (expected True: B's text should stay in Georgia, through a style or direct formatting)
```

### 4. A revision has no text (R05), and the CLI and Python spell kinds differently

`Revision` has `id`, `kind`, `author`, `timestamp`, `story`, not the text inserted, deleted or reformatted, and
`rdocx revision list --json` neither: listing tracked changes with their text needs `part_xml` and an XML parser
(the R05 implementation, 19 lines). Minor: `rdocx revision list --json` gives `run-property-change` where Python's
`Revision.kind` and `references/cli.md` say `run_property_change`.

Also seen, not counted as failures: a comment added by rdocx (R01) gets a `commentsExtended.xml` entry but no
`commentsIds.xml` entry and no `w:annotationRef` run; a reply (finding 1) no `w:pStyle CommentText` either.

## Findings on the other stacks

- Built-in: `comment.py` has no `--date`; there is no helper to resolve a thread, reject by author, redline or
  merge; pandoc does not report formatting changes; LibreOffice renders tracked changes unless they are accepted
  first, so the skill's own "render and look" loop shows the markup view; `accept_changes.py`'s fixed profile
  serialises every caller.
- python-docx 1.2: `add_comment` needs whole runs (no run split), writes no `w14:paraId` and no
  `commentsExtended.xml` entry (so no reply or resolution without lxml), and has no date setter; no API for
  revisions, fields, TOC, layout or rendering. A quirk met on the way: `itertext()` on a run element
  (`CT_R`) returns its text twice; iterating `w:t` children does not.
