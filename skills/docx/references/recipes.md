# docx recipes (each block is run by the test suite on the pinned build)

Conventions, as set up in SKILL.md: `R=${RDOCX_HOME:-~/.local/share/rdocx-skills}/current/bin`, `SKILL` is
this skill's folder, and Python blocks run with `PYTHONPATH="$SKILL/scripts" $R/python`, so that
`import docx_ops` works. `report.docx` is the input; outputs go to new files. Blocks run in order in one
folder: a later block may read a file an earlier one wrote.

## Read the text and find where things are

```bash
$R/rdocx text --json report.docx > report.json        # body and tables, then the other stories in "stories"
$R/python "$SKILL/scripts/docx_ops.py" text report.docx > all-stories.txt   # every paragraph with its part and style
$R/python "$SKILL/scripts/docx_ops.py" count report.docx "footbridge"
$R/rdocx layout --json report.docx > layout.json
```

Plain `rdocx text` shows the same accepted view (tracked insertions in, deletions out): the body, then each
other part (text boxes, headers, footers, notes, comments) under a `--- kind (part) ---` line.

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
stay out of reach of the replacement (a match across the edge of a content control, a tracked insertion or
a simple field). The message says where they are. `rdocx replace` does not make that second check. The command line equivalent:
`$R/python "$SKILL/scripts/docx_ops.py" replace IN OUT --edit OLD NEW COUNT --edit ...`.

## Replace in one paragraph or one table cell

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
# "high points of the deck" is in several paragraphs: change it in the method paragraph only
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("Drainage performance was checked"))
doc.paragraphs[i].replace_text("high points of the deck", "high points of each span", expect=1)
t = next(k for k, tb in enumerate(doc.tables) if tb.cell(0, 0).text == "Ref")
doc.tables[t].cell(1, 5).replace_text("No action", "Monitor", expect=1)   # F01's action, not the other rows'
docx_ops.save_atomic(doc, "scoped.docx", "report.docx")
```

A count other than `expect` raises `rdocx.ReplacementCountError` and changes nothing. The match may cross
runs, and a comment on the paragraph stays. `doc.replace_text_at(item, old, new, expect=N)` does
the same in one item of `doc.story_items` (a header, footer or footnote paragraph, a table cell's
paragraph); a text box or a comment is refused.

## Rewrite without losing or spreading run formatting

```python
import docx_ops, rdocx
doc = rdocx.Document()
p = doc.add_paragraph("")
p.add_run("Note:").font.bold = True                      # a bold run-in lead
p.add_run(" access to the soffit needs a permit.")
docx_ops.save_atomic(doc, "lead.docx")
doc = rdocx.Document.open("lead.docx")
doc.paragraphs[0].replace_text("a permit", "two permits", expect=1)   # anchored after the lead
docx_ops.save_atomic(doc, "lead-edited.docx", "lead.docx")
runs = lambda f: [[(r["text"], r["formatting"]) for r in q["runs"]] for q in docx_ops.textmap(f)]
before, after = runs("lead.docx"), runs("lead-edited.docx")
assert after[0][0] == before[0][0] and after[0][1][1] == before[0][1][1]   # the lead and the rest keep their format
```

`p.text = ...` leaves one unformatted run, and `set_story_text` gives the whole text the first run's format:
on this paragraph both lose the lead, one by dropping its bold, the other by making the whole paragraph bold. A
replacement whose old text starts inside the bold lead makes its new text bold too: anchor after the lead, or
edit `runs[k].text`. A placeholder filled in its own highlighted run keeps the highlight. `rdocx diff` compares
text only: compare the runs and their `formatting` (`rdocx text --json`, `docx_ops.textmap`) as well.

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
i, start, _ = docx_ops.locate(doc, "safety of path users")      # flow index, run offset, body index
first, last = docx_ops.isolate(doc, i, start, start + len("safety of path users"))
for k in range(first, last):
    doc.paragraphs[i].runs[k].font.bold = True
docx_ops.save_atomic(doc, "bold.docx", "report.docx")
```

## Style, numbering, spacing

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("The overall condition index"))
doc.paragraphs[i].style = "caption"                      # an id or a name, KeyError when the document has neither
doc.paragraphs[i].paragraph_format.space_after = rdocx.Pt(12)
doc.paragraphs[i].paragraph_format.keep_with_next = True
j = next(k for k, p in enumerate(doc.paragraphs) if p.numbering)
doc.paragraphs[j].numbering = (doc.paragraphs[j].numbering[0], 1)   # one level deeper
docx_ops.save_atomic(doc, "styled.docx", "report.docx")
```

## A heading numbered by its style

```python
import docx_ops, rdocx
doc = rdocx.Document()
steps = doc.add_numbering_instance(doc.add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.")]))
doc.link_style_to_numbering("Heading 1", steps, 0)          # every Heading 1 is numbered by its style
doc.add_paragraph("Scope")
doc.paragraphs[0].style = "Heading 1"
docx_ops.save_atomic(doc, "style-numbered.docx")
p = rdocx.Document.open("style-numbered.docx").paragraphs[0]
assert p.numbering is None and p.text == "Scope"           # the page shows "1. Scope"
```

`Paragraph.numbering` and the text views show only a paragraph's own numbering: a heading numbered by its
style reads unnumbered. Read the render before calling a numbering a defect; typing a number into its text
doubles it.

## A new list item, and the lead-in kept with its list

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
words = doc.word_count()
lead = next(k for k, p in enumerate(doc.paragraphs) if p.text == "The main findings are:")
last = lead + 1
while doc.paragraphs[last + 1].numbering == doc.paragraphs[lead + 1].numbering:
    last += 1
doc.clone_content(doc.paragraphs[last], doc.find_content_index(doc.paragraphs[last]) + 1)  # an item of this list
doc.paragraphs[last + 1].runs[0].text = "the lighting columns need repainting."
for k in range(1, len(doc.paragraphs[last + 1].runs)):
    doc.paragraphs[last + 1].runs[k].text = ""
doc.paragraphs[lead].paragraph_format.keep_with_next = True   # the lead-in never ends a page alone
docx_ops.save_atomic(doc, "list.docx", "report.docx")
new = rdocx.Document.open("list.docx")
assert new.paragraphs[last + 1].numbering == new.paragraphs[last].numbering
print(words, new.word_count())                               # count the words before and after
```

A new item takes the paragraph properties of an existing item of the same list (clone it), so its numbering,
indents and spacing match. To split a highlighted paragraph, cut it run by run (`runs[k].text`) so that the
highlight survives.

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

## Comments: on exact text, on a table cell, reply, resolve

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
cid = docx_ops.comment_on_text(doc, "about 12 mm", "Measured against which reference?", "Reviewer", "RV")
rid = doc.reply_to(cid, author="Author", text="The 2019 survey marks.", date=docx_ops.now())
doc.resolve_comment(cid)
docx_ops.comment_on_text(doc, "spalled concrete", "Which face of the pier?", "Reviewer", in_tables=True)  # part of a cell
cell = next(it for it in doc.story_items if it.story.kind == "table_cell" and it.kind == "paragraph" and it.text == "Ref")
doc.add_comment(rdocx.StoryRunRange(start=rdocx.StoryRunPosition(item=cell, run_index=0),
                                    end=rdocx.StoryRunPosition(item=cell, run_index=len(doc.tables[0].cell(0, 0).paragraphs[0].runs))),
                author="Reviewer", text="Rename this column?", date=docx_ops.now())
docx_ops.save_atomic(doc, "commented.docx", "report.docx")
```

`comment_on_text` splits runs so that the comment covers exactly the anchor, also inside or after a content
control or a tracked insertion, dates it, and refuses (nothing changed) if rdocx would anchor it anywhere
else, or if the anchor sits inside a simple field, a smart tag or a custom XML element. `occurrence` counts
the body's own paragraphs; with `in_tables=True` (`--in-tables`) it counts the paragraphs of table cells,
nested tables included, and the comment covers exactly the anchor inside the cell. `add_comment` without
`date=` writes an undated comment.

```bash
$R/rdocx comment list --json commented.docx
$R/python "$SKILL/scripts/docx_ops.py" comment report.docx commented-cli.docx --anchor "about 12 mm" --text "Which reference?" --author "Reviewer"
$R/python "$SKILL/scripts/docx_ops.py" comment report.docx commented-cell.docx --anchor "spalled concrete" --in-tables --text "Which face?" --author "Reviewer"
```

## Review pass: what each comment is on, move a thread, delete a paragraph

```python
import docx_ops, rdocx
doc = rdocx.Document.open("commented.docx")
for c in doc.comments:                                   # a reply has no anchor of its own (None)
    print(c.id, c.parent_id, c.author, repr(c.anchor_text), "->", c.text)
# the bearing item is to go; its thread is about the measuring method: move it there first
cid = next(c.id for c in doc.comments if c.parent_id is None and c.anchor_text == "about 12 mm")
doc.move_comment_to_text(cid, "bearing positions")      # id, replies and resolved flag kept
i = next(k for k, p in enumerate(doc.paragraphs) if p.text.startswith("the bearing at pier 2"))
doc.remove_content(doc.find_content_index(doc.paragraphs[i]))
docx_ops.save_atomic(doc, "reviewed.docx", "commented.docx")
```

```bash
$R/rdocx comment list --json reviewed.docx                # anchor_text and anchor of each comment
$R/rdocx comment move commented.docx 0 --text "bearing positions" -o moved-cli.docx --json
$R/rdocx validate reviewed.docx                           # exit 1 for a comment left without an anchor
```

Removing content never leaves a comment pointing at nothing: a comment the removed block covers whole goes
with its replies (move it first to keep it). A removal that would cut a comment in part, and `pop_content` of
a block that carries a comment, raise `RdocxError` and change nothing: move the thread off the block first, or
move the block with `move_content`, which keeps its threads. `Table.remove_row` removes a comment the row
covers whole the same way.

## Redline two versions, then accept or reject

```bash
cp report.docx v1.docx
$R/rdocx replace v1.docx -p "three points lower" -v "two points lower" --expect 1 -o v2.docx
$R/rdocx compare v1.docx v2.docx --author "Reviewer" --timestamp 2026-09-27T12:00:00Z --granularity word -o redline.docx --json
$R/rdocx revision list --json redline.docx
$R/rdocx revision accept redline.docx --author "Reviewer" -o accepted.docx --json
$R/rdocx convert redline.docx --to pdf --revision-view tracked -o redline-tracked.pdf
```

The redline holds only the edit, in every story (`revision list` shows each revision's story), and
`--granularity word` marks only the changed words (the default, `run`, deletes and re-inserts the whole
run). In Python: `v1.compare(v2, "Reviewer", timestamp, granularity="word")`. Compare before rebuilding the
TOC: a TOC rebuilt on the edited side adds its entries to the redline as revisions.
To show the redline outside Word, render it with `--revision-view tracked` (Python
`to_pdf(revision_view="tracked")`): deletions struck through, insertions underlined, a change bar. The default
PDF is the accepted view.

## Table of contents and page fields

```bash
$R/python "$SKILL/scripts/docx_ops.py" toc report.docx toc.docx
```

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
rep = doc.rebuild_toc()
fields = doc.update_layout_backed_fields()
docx_ops.save_atomic(doc, "fields.docx", "report.docx")
print(rep.entry_count, fields.updated_count)
```

After a save in Google Docs or Word (the fixture's footer fields are packed with no cached result), refresh the
fields and stop Word from asking to update them on open:

```python
import docx_ops, rdocx, zipfile
doc = rdocx.Document.open("report.docx")
doc.update_layout_backed_fields()
doc.update_fields_on_open = False
docx_ops.save_atomic(doc, "no-prompt.docx", "report.docx")
with zipfile.ZipFile("no-prompt.docx") as z:
    assert b'<w:updateFields w:val="false"/>' in z.read("word/settings.xml")
```

A document without a table of contents gets one with `insert_toc`, then `rebuild_toc` fills it:

```python
import docx_ops, rdocx
doc = rdocx.Document()
doc.add_paragraph("Inspection report")
doc.paragraphs[0].style = "Title"
for heading in ("Scope", "Findings", "Recommendations"):
    doc.add_paragraph(heading)
    doc.paragraphs[-1].style = "Heading 1"
    doc.add_paragraph("Body text. " * 300)
doc.insert_toc(1, max_level=2)                 # a TOC field at body index 1, after the title
rep = doc.rebuild_toc()                        # entries, links and page numbers from rdocx's pagination
docx_ops.save_atomic(doc, "with-toc.docx")
print(rep.entry_count)
```

## Check what you changed

```bash
$R/python "$SKILL/scripts/docx_ops.py" replace report.docx checked.docx --edit "three points lower" "two points lower" 1
$R/rdocx validate checked.docx
$R/rdocx diff report.docx checked.docx
pages=$(mktemp -d)                                       # render into a new folder each time
$R/rdocx render checked.docx -o "$pages" --pages 1-3 --dpi 80
ls "$pages"
```

`validate` reads every part the document relates to and checks every style id. Look at the PNG of every
page you touched. Page numbers come from rdocx's layout, which
lays lines at Word's heights: still check in Word before quoting one as Word's.

## Page breaks before cutting text

```python
import docx_ops, rdocx
doc = rdocx.Document.open("report.docx")
breaks = [i for i, p in enumerate(doc.paragraphs) if p.paragraph_format.page_break_before]   # active ones only
print(doc.page_count(), [(i, doc.paragraphs[i].text[:40]) for i in breaks])
for i in breaks:
    variant = rdocx.Document.open("report.docx")
    variant.paragraphs[i].paragraph_format.page_break_before = False
    docx_ops.save_atomic(variant, f"no-break-{i}.docx", "report.docx")
    print(i, variant.page_count())          # then render the variant and look at the pages around the break
```

Google Docs writes `pageBreakBefore` on every paragraph, mostly as `w:val="0"`: `page_break_before` is True
only for an active break (None or False otherwise). Whether a page goes away depends on the layout around the
break: render each variant before cutting any text.

## Before deleting a package part

```python
import posixpath, zipfile
part = "word/media/image1.png"
with zipfile.ZipFile("report.docx") as z:
    users = [n for n in z.namelist() if n.endswith(".rels") and posixpath.basename(part).encode() in z.read(n)]
print(part, users)                          # an empty list only: no part relates to it
```

A part is not orphan until its name has been searched in every `.rels` of the package, not only the
document's: an image can be used by the theme, a header or a chart alone.

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
From a .dotx, save as .docx the same way: rdocx writes the content type the extension names.

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
with zipfile.ZipFile("from-template.docx") as z:
    assert b"wordprocessingml.document.main+xml" in z.read("[Content_Types].xml")
```

## A new document with its own styles, a numbered list and properties

```python
import docx_ops, rdocx
doc = rdocx.Document()                                     # Normal, Title, Heading 1 to 9, List Paragraph...
doc.core_properties.title = "Inspection checklist"
doc.core_properties.author = "Claude"
note = doc.add_style("Note box", based_on="Normal", italic=True, left_indent=rdocx.Inches(0.5))
doc.set_style("Normal", font_size=rdocx.Pt(11), space_after=rdocx.Pt(6))  # an existing style, the rest kept
steps = doc.add_numbering_instance(doc.add_numbering_definition([
    rdocx.ListLevel(format="decimal", text="%1.", left_indent=rdocx.Inches(0.5), hanging_indent=rdocx.Inches(0.25))]))
doc.add_paragraph("Inspection checklist")
doc.paragraphs[0].style = "Title"
doc.add_paragraph("Bring the 2019 survey marks.")
doc.paragraphs[1].style = note.style_id
for text in ("Check the bearings.", "Photograph the deck joints."):
    doc.add_paragraph(text)
    doc.paragraphs[-1].style = "List Paragraph"
    doc.paragraphs[-1].numbering = (steps, 0)
doc.paragraphs[1].text = "Bring the 2019 survey marks and a tape."    # one plain run, paragraph style kept
docx_ops.save_atomic(doc, "checklist.docx")
```

`add_style` derives the id from the name as Word does (`"Note box"` → `Notebox`). Assign by that id or by
the name. `set_style` changes only the properties it is given, on a style chosen by id or name.
`link_style_to_numbering(style, num_id, level)` numbers every paragraph of a style instead.
