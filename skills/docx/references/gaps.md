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
| Validation [validate-parts] | `rdocx validate` passes a truncated header or footer and a paragraph whose style id the package does not define | verify with `validate`, then `Document.open(out)` and `doc.story_items`, then a render |

## Editing

| Gap | What happens | Workaround | Fallback |
|---|---|---|---|
| Tracked insertions [replace-tracked-insertions] | replacement does not see text inside `w:ins`, although `Paragraph.text` and `text --json` show it | accept or reject the revisions first when that is the intent; `docx_ops.replace_batch` refuses otherwise | lxml, one targeted edit |
| Footnotes and endnotes [replace-footnotes] | replacement skips them, while headers, footers, cells and text boxes are reached | `docx_ops.replace_batch` refuses; `docx_ops.count` shows the notes | lxml on `word/footnotes.xml` / `word/endnotes.xml` |
| No paragraph text setter [docx-paragraph-text-setter] | `Paragraph.text` is read-only | `doc.set_story_text(item, text)` (keeps the first run's format; an empty run stays), or set `runs[0].text` and blank the others | none needed |
| Unchecked style ids [style-id-unchecked] | `Paragraph.style = "Heading 1"` (a name) or an unknown id is stored as is; `numbering` is not checked either | check against `{s.style_id for s in doc.styles}` before assigning | none needed |
| Styles and numbering [docx-python-styles, new-document-styles] | cannot create a style or a list definition; a new `Document()` has only Normal and Heading1 | start from a template .docx that defines them (a copy of a .docx, see the template gap), and assign by style id | python-docx on a template |
| Core properties [docx-core-properties] | no title, author, subject or dates, read or write; `validate` warns "Missing document title" on every new document | keep the template's properties; `rdocx inspect --json` reads title, author, subject, keywords | python-docx `core_properties` for that step |
| Template content type [template-save-as-document] | a .dotx opened and saved as .docx keeps the template content type (Word expects a document; not checked in Word) | after the save, rewrite the main part's content type in `[Content_Types].xml` (`template.main+xml` → `document.main+xml`): recipe "A document from a .dotx template" | none needed (python-docx refuses to open a .dotx) |
| Hyperlinks [docx-hyperlink-retarget] | no API to change a link's target or to unlink its text | `set_story_text` rewrites the paragraph without its links (their relationships stay in the package, unused), then `add_hyperlink` for the new target | lxml |
| Picture size [docx-picture-resize] | cannot resize an existing picture; `replace_image` keeps the old extent, so a new aspect ratio renders distorted | give the new picture the old pixel aspect ratio | lxml on `wp:extent` and `a:ext` |

## Review: comments, tracked changes, redline

| Gap | What happens | Workaround / fallback |
|---|---|---|
| CLI comment dates [comment-date-cli] | `comment add` and `comment reply` have no `--date`: CLI comments are undated | `docx_ops.py comment ... --date`, or `add_comment(..., date=)` in Python |
| Whole-run redline in the CLI [compare-granularity] | `rdocx compare` marks a whole run as deleted and re-inserted for one changed word; single-run paragraphs (Google Docs exports, typed text) read as rewritten | Python `a.compare(b, author, timestamp, granularity="word")` marks only the changed words |
| Refusals [compare-comments, compare-rebuilt-toc] | `compare()` and `rdocx compare` refuse the pair when comments differ (only Python can ignore them) or when the edited side had its TOC rebuilt | Python `compare(..., ignore_comments=True)` keeps the original's comments and compares the rest; compare before rebuilding the TOC; if it still refuses, report the redline as unavailable. Fallback for a human-facing redline: Word's Compare, done by the user |

## Fields, layout, rendering

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Line gap [line-gap] | a single line is laid out without the font's line gap: Calibri 1.000 em where Word uses 1.221 em, Arial 1.117 against 1.150 | page counts and page breaks differ from Word, most with Calibri: never state a Word page number from rdocx's layout; check in Word when it matters |
| Picture lines [picture-line-spacing] | a line holding an inline picture is multiplied by proportional spacing (400 pt picture at line 264: 440 pt; Word 402.7 pt) | tall figures can be split from their caption in rdocx's render only |
| Tab stops [tab-stops] | text after a tab stop set in the paragraph or its style starts 36 pt (half an inch) before the stop, and a right-aligned stop is laid out as a left one: a page number after a dot leader ends short of its stop by 36 pt minus its own width | rdocx's render only; the file is untouched. Never judge tab alignment (TOC, forms, signature lines) from rdocx's PDF or PNG; check it in the target application |
| Numbered TOC entries [toc-numbered-entries] | for a numbered heading, `rebuild_toc()` writes number, tab, title, tab, page and adds no stop for the first tab: when the TOC style has only its right page-number stop, the title is pushed to the right margin | before rebuilding, give the TOC styles a left tab stop after the number (python-docx, one step on the styles), or leave the TOC to Word (`doc.update_fields_on_open = True`); rdocx's own render of the entries stays off (gap tab-stops) |
