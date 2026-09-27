# docx recipes (each block is run by the test suite on the pinned build)

Conventions, as set up in SKILL.md: `R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, `SKILL` is
this skill's folder, and Python blocks run with `PYTHONPATH="$SKILL/scripts" $R/python`, so that
`import docx_ops` works. `report.docx` is the input; outputs go to new files. Blocks run in order in one
folder: a later block may read a file an earlier one wrote.

## Read the text and find where things are

```bash
$R/rdocx text --json report.docx > report.json        # body and tables, accepted view of tracked changes
$R/python "$SKILL/scripts/docx_ops.py" text report.docx > all-stories.txt   # plus headers, footers, notes, text boxes
$R/python "$SKILL/scripts/docx_ops.py" count report.docx "footbridge"
$R/rdocx layout --json report.docx > layout.json
```

Plain `rdocx text` is fine for a quick look at a clean file; it drops tracked insertions and content-control
blocks, so do not use it on a document under review.

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
# flow index, style and text of every heading
heads = [(i, p.style, p.text) for i, p in enumerate(doc.paragraphs) if (p.style or "").startswith("Heading")]
# the body index of one paragraph, and the page it starts on
i = next(i for i, _, t in heads if t == "Recommendations")
bi = doc.find_content_index(doc.paragraphs[i])
page = next(f.physical_page for f in doc.layout() if f.body_index == bi)
# every paragraph with its path, table cells included
cells = [p for p in docx_ops.textmap("report.docx") if len(p["path"]) > 1]
# headers and footers, one story per variant part
footers = [t for kind, t in docx_ops.all_text(doc) if kind == "footer"]
print(len(heads), bi, page, len(cells), footers)
```

## Counted replacements, all or nothing

```bash
$R/rdocx replace report.docx -p "Issue B" -v "Issue C" --expect 2 -o issue-c.docx
```

```python
import docx_ops
docx_ops.replace_batch("report.docx", "edited.docx", [
    ("Issue B, for review", "Issue C, for approval", 1),
    ("three points lower", "two points lower", 1),
])
```

`replace_batch` raises `docx_ops.EditError`, and writes nothing, when a count differs or when occurrences
stay out of reach of the replacement (content controls, tracked insertions, footnotes, endnotes); the
message says where they are. `rdocx replace` does not make that second check. The command line equivalent:
`$R/python "$SKILL/scripts/docx_ops.py" replace IN OUT --edit OLD NEW COUNT --edit ...`.

## A new paragraph after an anchor, with the anchor's formatting

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("None of these defects"))
bi = doc.find_content_index(doc.paragraphs[i])
doc.clone_content(doc.paragraphs[i], bi + 1)
new = i + 1                                   # flow index of the clone
doc.paragraphs[new].runs[0].text = "The next principal inspection is due in 2031."
for k in range(1, len(doc.paragraphs[new].runs)):
    doc.paragraphs[new].runs[k].text = ""
docx_ops.save_atomic(doc, "added.docx", "report.docx")
```

Clone only a plain paragraph: a clone copies fields and hyperlinks, and bookmarks under a new name
(`MailMerge1`...); comment anchors are not copied.

## Delete or move a block

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("Access equipment, traffic management"))
doc.remove_content(doc.find_content_index(doc.paragraphs[i]))
j = next(k for k, p in enumerate(doc.paragraphs) if p.text == "The main findings are:")
doc.move_content(doc.paragraphs[j], doc.find_content_index(doc.paragraphs[j]) + 2)   # counted before the move
docx_ops.save_atomic(doc, "restructured.docx", "report.docx")
```

## Bold one phrase inside a run

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
i, start, _ = docx_ops.locate(doc, "safety of path users")      # flow index, offset, body index
first, last = docx_ops.isolate(doc, i, start, start + len("safety of path users"))
for k in range(first, last):
    doc.paragraphs[i].runs[k].font.bold = True
docx_ops.save_atomic(doc, "bold.docx", "report.docx")
```

## Style, numbering, spacing

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
ids = {s.style_id for s in doc.styles}
assert "Caption" in ids                                  # rdocx stores any string: check the id first
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("The overall condition index"))
doc.paragraphs[i].style = "Caption"
doc.paragraphs[i].paragraph_format.space_after = rdocx.Pt(12)
doc.paragraphs[i].paragraph_format.keep_with_next = True
j = next(k for k, p in enumerate(doc.paragraphs) if p.numbering)
doc.paragraphs[j].numbering = (doc.paragraphs[j].numbering[0], 1)   # one level deeper
docx_ops.save_atomic(doc, "styled.docx", "report.docx")
```

## Tables

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
t = next(k for k, tb in enumerate(doc.tables) if tb.cell(0, 0).text == "Item")
n = len(doc.tables[t].rows)
doc.tables[t].clone_row(n - 2)                            # copy the last item row, formatting included
new = n - 1                                               # the clone sits after its source
for c, text in enumerate(["Replace the handrail caps", "8", "no.", "40", "320"]):
    doc.tables[t].cell(new, c).text = text
doc.tables[t].rows[n].cells[-1].text = "13,242"           # the total row moved down by one (merged cells: 2)
docx_ops.save_atomic(doc, "table.docx", "report.docx")
```

## Replace a picture, add a picture

```python
import re, docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
drawing = next(it for it in doc.story_items if it.kind == "drawing")
xml = drawing.xml.decode() if isinstance(drawing.xml, bytes) else drawing.xml
rid = re.search(r'r:embed="([^"]+)"', xml).group(1)
w, h = docx_ops.image_size(doc.image_data(rid))
new = docx_ops.solid_png(w, h, (40, 90, 140))   # stands for the new picture's bytes; keep the aspect ratio
doc.replace_image(rid, new)
docx_ops.save_atomic(doc, "picture-replaced.docx", "report.docx")

note = rdocx.Document()
note.add_paragraph("Figure A1")
note.add_picture(new, "a1.png", width=rdocx.Inches(3), height=rdocx.Inches(3 * h / w))
docx_ops.save_atomic(note, "picture-added.docx")
```

`add_picture` fails on a file that has a content control and a default namespace on its root, as Google Docs
exports do (gap add-picture-sdt-default-ns): for those, fall back to python-docx for that step.

## Comments: on exact text, on a table cell, reply, resolve

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
cid = docx_ops.comment_on_text(doc, "about 12 mm", "Measured against which reference?", "Reviewer", "RV")
rid = doc.reply_to(cid, author="Author", text="The 2019 survey marks.", date=docx_ops.now())
doc.resolve_comment(cid)
cell = next(it for it in doc.story_items if it.story.kind == "table_cell" and it.kind == "paragraph" and it.text == "Ref")
doc.add_comment(rdocx.StoryRunRange(start=rdocx.StoryRunPosition(item=cell, run_index=0),
                                    end=rdocx.StoryRunPosition(item=cell, run_index=len(doc.tables[0].cell(0, 0).paragraphs[0].runs))),
                author="Reviewer", text="Rename this column?", date=docx_ops.now())
docx_ops.save_atomic(doc, "commented.docx", "report.docx")
```

`comment_on_text` splits runs so that the comment covers exactly the anchor, dates it, and refuses (nothing
changed) when the paragraph holds a content control or a tracked insertion before the anchor (gap
comment-runposition-sdt). `add_comment` without `date=` writes an undated comment.

```bash
$R/rdocx comment list --json commented.docx
$R/python "$SKILL/scripts/docx_ops.py" comment report.docx commented-cli.docx --anchor "about 12 mm" --text "Which reference?" --author "Reviewer"
```

## Redline two versions, then accept or reject

```bash
$R/python -c 'import rdocx; rdocx.Document.open("report.docx").save("v1.docx")'
$R/rdocx replace v1.docx -p "three points lower" -v "two points lower" --expect 1 -o v2.docx
$R/rdocx compare v1.docx v2.docx --author "Reviewer" --timestamp 2026-09-27T12:00:00Z -o redline.docx --json
$R/rdocx revision list --json redline.docx
$R/rdocx revision accept redline.docx --author "Reviewer" -o accepted.docx --json
```

The first line passes the original through rdocx: a file with an empty comments part (every Google Docs
export) is otherwise refused against its rdocx-edited version (gap empty-comments-reserialised). Compare
first, then rebuild the TOC or refresh fields: the other order is refused. A one-word change shows as its
whole run deleted and re-inserted, and the redline may carry a section property change with no visible
difference (gap compare-own-save-noise). `revision list` shows the main story only; the `resolved` count of
`revision accept --json` covers headers, footers and notes too.

## Table of contents and page fields

```bash
$R/python "$SKILL/scripts/docx_ops.py" toc report.docx toc.docx
```

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
rep = docx_ops.rebuild_toc(doc)
fields = doc.update_layout_backed_fields()
docx_ops.save_atomic(doc, "fields.docx", "report.docx")
print(rep.entry_count, fields.updated_count)
```

## Check what you changed

```bash
$R/python "$SKILL/scripts/docx_ops.py" replace report.docx checked.docx --edit "three points lower" "two points lower" 1
$R/rdocx validate checked.docx
$R/python -c 'import rdocx; d = rdocx.Document.open("checked.docx"); d.story_items; print(len(d.paragraphs))'
$R/rdocx diff report.docx checked.docx
pages=$(mktemp -d)                                       # render into a new folder each time
$R/rdocx render checked.docx -o "$pages" --pages 1-3 --dpi 80
ls "$pages"
```

`validate` misses broken headers and dangling style ids (gap validate-parts): the re-open with `story_items`
reads every story. Look at the PNG of every page you touched. Page numbers come from rdocx's layout, which
differs from Word's (gaps line-gap, picture-line-spacing): never quote them as Word's.

## A new document from a template

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")                  # any .docx that defines the styles you need
while len(doc.paragraphs) or len(doc.tables):
    doc.remove_content(0)
doc.add_paragraph("Site visit note")
doc.paragraphs[0].style = "Heading1"
doc.add_paragraph("Visited on 27 September 2026.")
docx_ops.save_atomic(doc, "note.docx")
```

`remove_content` leaves the section properties, headers, footers, styles and numbering of the template.
From a .dotx, save as .docx the same way, then fix the content type that rdocx keeps (gap
template-save-as-document):

```python
import docx_ops, rdocx, zipfile
with zipfile.ZipFile("note.docx") as src, zipfile.ZipFile("note.dotx", "w") as dst:   # a .dotx for this example
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename == "[Content_Types].xml":
            data = data.replace(b"document.main+xml", b"template.main+xml")
        dst.writestr(info, data)
doc = rdocx.Document.open("note.dotx")
doc.add_paragraph("Filled from the template.")
docx_ops.save_atomic(doc, "from-template.docx")
docx_ops.fix_template_content_type("from-template.docx")
with zipfile.ZipFile("from-template.docx") as z:
    assert b"wordprocessingml.document.main+xml" in z.read("[Content_Types].xml")
```
