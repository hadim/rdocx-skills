# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

## Reading

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Other stories [text-other-stories] | `rdocx text`, `text --json` and `convert --to md/html` hold the body and its tables only: no headers, footers, footnotes, endnotes, text boxes | `docx_ops.py text F` (reads the package), or `doc.story_items` by `item.story.kind` |
| Word text boxes [textbox-alternate-content] | a text box as Word writes it (`mc:AlternateContent`: DrawingML with a VML copy) is missing from `doc.story_items`, and replacement counts it twice | `docx_ops.story_paragraphs` and `count` read it once; expect rdocx's count to include the copy |
| Simple fields, smart tags [text-wrapped-runs] | text inside `w:fldSimple`, `w:smartTag` and inline `w:customXml` is missing from `Paragraph.text` and `text --json`, and replacement does not see it | `docx_ops.story_paragraphs` reads it; `docx_ops.replace_batch` refuses when it is left; fallback: lxml, one targeted edit |
| Tracked insertions, plain text [text-cli-tracked-insertions] | plain `rdocx text` drops the text of tracked insertions: it shows neither the accepted nor the original version | `rdocx text --json` (accepted view) or `Paragraph.text` |
| Body-level content controls, plain text [sdt-text-cli] | plain `rdocx text` skips body-level content controls | `rdocx text --json`, `doc.paragraphs` |
| Validation [validate-parts] | `rdocx validate` passes a truncated header or footer and a paragraph whose style id the package does not define | verify with `validate`, then `Document.open(out)` and `doc.story_items`, then a render |

## Editing

| Gap | What happens | Workaround | Fallback |
|---|---|---|---|
| Content controls [sdt-replace] | `try_replace_text`, `replace_all_regex` and `rdocx replace` do not see text inside a `w:sdt` (Google Docs wraps text in `goog_rdk_*` controls; Word uses them for forms and the TOC) | `docx_ops.replace_batch` refuses when occurrences stay out of reach; `docx_ops.count` and `doc.paragraphs` show where they are. A TOC is rebuilt after the edit (`docx_ops.rebuild_toc`) | lxml on `word/document.xml`, one targeted edit |
| Tracked insertions [replace-tracked-insertions] | replacement does not see text inside `w:ins`, although `Paragraph.text` and `text --json` show it | accept or reject the revisions first when that is the intent; `docx_ops.replace_batch` refuses otherwise | lxml, one targeted edit |
| Footnotes and endnotes [replace-footnotes] | replacement skips them, while headers, footers, cells and text boxes are reached | `docx_ops.replace_batch` refuses; `docx_ops.count` shows the notes | lxml on `word/footnotes.xml` / `word/endnotes.xml` |
| Tables in headers and footers [replace-header-footer-tables] | replacement skips the cells of a table inside a header or footer | `docx_ops.replace_batch` refuses and names the part | lxml on that `word/footerN.xml`, one targeted edit |
| No paragraph text setter [docx-paragraph-text-setter] | `Paragraph.text` is read-only | `doc.set_story_text(item, text)` (keeps the first run's format; an empty run stays), or set `runs[0].text` and blank the others | none needed |
| `split_run` index [split-run-index] | `split_run(i, ...)` counts paragraphs only while `RunPosition` counts body children: after a table, a split lands on another paragraph | pass the `doc.paragraphs` index to `split_run` and the body index to `RunPosition`; `docx_ops.comment_on_text` does it | none needed |
| Unchecked style ids [style-id-unchecked] | `Paragraph.style = "Heading 1"` (a name) or an unknown id is stored as is; `numbering` is not checked either | check against `{s.style_id for s in doc.styles}` before assigning | none needed |
| Tables in Python [docx-python-tables] | no merge, borders, shading, margins, grid, row height, header row from Python (they exist in Rust); cell and table widths can be set | edit a template table that already has them, then `clone_row` | python-docx for the table step |
| Styles and numbering [docx-python-styles, new-document-styles] | cannot create a style or a list definition; a new `Document()` has only Normal and Heading1 | start from a template .docx that defines them (a copy of a .docx, see the template gap), and assign by style id | python-docx on a template |
| Sections [docx-python-sections] | `Section` is read-only (margins, orientation, size, columns) | template file | python-docx for the section step |
| Bookmarks and fields [docx-python-bookmarks] | no bookmark or field insertion from Python | | lxml |
| Core properties [docx-core-properties] | no title, author, subject or dates, read or write; `validate` warns "Missing document title" on every new document | keep the template's properties; `rdocx inspect --json` reads title, author, subject, keywords | python-docx `core_properties` for that step |
| Template content type [template-save-as-document] | a .dotx opened and saved as .docx keeps the template content type (Word expects a document; not checked in Word) | after the save, rewrite the main part's content type in `[Content_Types].xml` (`template.main+xml` → `document.main+xml`): recipe "A document from a .dotx template" | none needed (python-docx refuses to open a .dotx) |
| Hyperlinks [docx-hyperlink-retarget] | no API to change a link's target or to unlink its text | `set_story_text` rewrites the paragraph without its links (their relationships stay in the package, unused), then `add_hyperlink` for the new target | lxml |
| Adding a picture to a Google Docs file [add-picture-sdt-default-ns] | `add_picture` raises "cannot serialize a modified document with a shadowed default namespace" when the body holds a content control and the root declares a default namespace (Google Docs exports) | none in rdocx | python-docx for the picture step |
| Picture size [docx-picture-resize] | cannot resize an existing picture; `replace_image` keeps the old extent, so a new aspect ratio renders distorted | give the new picture the old pixel aspect ratio | lxml on `wp:extent` and `a:ext` |
| Counted replacement contract in Python [docx-replace-contract] | `try_replace_text` returns a count but has no expected count and is not all-or-nothing | `docx_ops.replace_batch` (dry run on a copy, then apply) | none needed |
| Table row identity [tr-identity-lost] | any modelled edit drops `w:rsidR`, `w:rsidTr`, `w14:paraId` from every `w:tr` | harmless for Word and Google Docs; it shows in XML diffs | none needed |

## Saving and the CLI

| Gap | What happens | Workaround |
|---|---|---|
| Non-atomic save [save-not-atomic] | `Document.save(path)` writes the target in place: an interruption leaves a truncated file | `docx_ops.save_atomic(doc, out, src)` |
| `convert` overwrites [cli-convert-overwrites] | `rdocx convert --to pdf/md/html` overwrites any existing output without a word, the input included (`-o F` on the input, or the default output next to it) | always give `convert` a new path; the editing commands, `render` and image `convert` refuse an existing output |
| Empty comments part [empty-comments-reserialised, ignorable-undeclared] | a save rewrites an empty `word/comments.xml` (Google Docs exports have one); if its root lists `w14` in `mc:Ignorable`, the saved part leaves it undeclared, which strict consumers reject | check with `rdocx validate` and open the result in the target application when the file carries an empty comments part |
| Root `mc:Ignorable` dropped [ignorable-dropped] | an edit that rewrites `word/document.xml`, a header or a footer (a replacement in that story, for example) drops `mc:Ignorable` from the part's root, while its namespace declarations and the `w14:paraId` attributes stay: a consumer that does not know `w14` may then reject the part | when the file goes to a strict consumer or validator, copy the source root's `mc:Ignorable` back onto each rewritten part (a zip rewrite of that one attribute) |
| Broken pipe [cli-broken-pipe] | `rdocx text F \| head` panics (exit 101) when the reader closes the pipe | write to a file, or ignore exit 101 in that pattern |

## Review: comments, tracked changes, redline

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Comments beside content controls or insertions [comment-runposition-sdt] | the run index of `RunPosition` and `rdocx comment add --start-run/--end-run` skips runs inside an inline content control or a tracked insertion, while `Paragraph.runs` and `text --json` count them: the comment lands on the wrong text, no error | `docx_ops.comment_on_text` checks the anchored text and refuses; anchor on text before the control, or on the whole paragraph. Fallback: python-docx cannot add comments either; say that the comment could not be placed |
| CLI comment dates [comment-date-cli] | `comment add` and `comment reply` have no `--date`: CLI comments are undated | `docx_ops.py comment ... --date`, or `add_comment(..., date=)` in Python |
| Whole-run redline [compare-granularity] | `compare()` marks a whole run as deleted and re-inserted for one changed word; single-run paragraphs (Google Docs exports, typed text) read as rewritten | acceptable for a machine check; for a human reviewer, say that the redline is coarse |
| Refusals [compare-comments, compare-sdt-id, compare-rebuilt-toc, compare-packed-fields, empty-comments-reserialised] | `compare()` refuses the pair when: comments differ; a content control differs only by `w:id` or `w:tag`; the edited side had its TOC rebuilt; page fields packed in one run were refreshed; the original has an empty comments part and the edited file was saved by rdocx | compare before rebuilding the TOC or refreshing fields; for the empty comments part, pass the original through rdocx first (open and save), then compare; if it still refuses, report the redline as unavailable. Fallback for a human-facing redline: Word's Compare, done by the user |
| Serialisation noise [compare-own-save-noise] | a file compared with its own rdocx-edited copy can show a `section_property_change` (`w:orient="portrait"` dropped) or a deleted and re-inserted run (`xml:space` dropped) | ignore revisions whose text is unchanged; say so when you present the redline |
| Revisions outside the body [revisions-main-story] | `doc.revisions` and `rdocx revision list` list the main story only, while accept and reject act on headers, footers and notes too | `rdocx revision accept F -o TMP --json`: its `resolved` count covers every story |

## Fields, layout, rendering

| Gap | What happens | Workaround / fallback |
|---|---|---|
| TOC after Word [toc-rsid-field-runs] | `rebuild_toc()` fails on a fresh open (`root attribute prefix w is unbound`) when a run of the TOC field carries `w:rsid*`, as Word writes them | `docx_ops.rebuild_toc(doc)` (or `docx_ops.py toc IN OUT`): replaces the text of one unique run by itself, which is a modelled edit, then rebuilds; a zero-count replacement or a field refresh is not enough |
| Line gap [line-gap] | a single line is laid out without the font's line gap: Calibri 1.000 em where Word uses 1.221 em, Arial 1.117 against 1.150 | page counts and page breaks differ from Word, most with Calibri: never state a Word page number from rdocx's layout; check in Word when it matters |
| Picture lines [picture-line-spacing] | a line holding an inline picture is multiplied by proportional spacing (400 pt picture at line 264: 440 pt; Word 402.7 pt) | tall figures can be split from their caption in rdocx's render only |
| Tab stops [tab-stops] | text after a tab stop set in the paragraph or its style starts 36 pt (half an inch) before the stop, and a right-aligned stop is laid out as a left one: a page number after a dot leader ends short of its stop by 36 pt minus its own width | rdocx's render only; the file is untouched. Never judge tab alignment (TOC, forms, signature lines) from rdocx's PDF or PNG; check it in the target application |
| Numbered TOC entries [toc-numbered-entries] | for a numbered heading, `rebuild_toc()` writes number, tab, title, tab, page and adds no stop for the first tab: when the TOC style has only its right page-number stop, the title is pushed to the right margin | before rebuilding, give the TOC styles a left tab stop after the number (python-docx, one step on the styles), or leave the TOC to Word (`doc.update_fields_on_open = True`); rdocx's own render of the entries stays off (gap tab-stops) |
| PDF text layer [pdf-text-ligatures] | with Calibri (Carlito) text, the PDF's text layer (copy, search, `pdftotext`) is garbled after the first ligature (`ti`, `fi`, `ff`); the page itself renders correctly | verify rendered text with `rdocx text` on the .docx, not `pdftotext` on the PDF; for a PDF whose text must be searchable, fallback to LibreOffice for the export |
