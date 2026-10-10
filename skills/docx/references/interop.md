# Files that travel between Word, Google Docs and LibreOffice

A .docx is rarely written by one application only. Each producer writes the same content its own way, and
the next one must still open what you save. What matters when you edit with rdocx:

## What each producer leaves in the XML

| Trait | Written by | What to do |
|---|---|---|
| `w:rsid*` on paragraphs, runs and rows, `w14:paraId` / `w14:textId`, a `w:rsids` list in settings | Word, on every save | carry no content; rdocx keeps them, on table rows too |
| explicit `w:val="0"` toggles (`w:b`, `w:i`, `w:rtl`, `keepNext`, `pageBreakBefore`), `xml:space="preserve"` on every `w:t`, `w:orient="portrait"` | Google Docs | equivalent to the defaults; a redline against an rdocx-edited copy shows only the edit |
| content controls `w:sdt` with `w:tag goog_rdk_N` around runs or paragraphs | Google Docs | read, replaced and commented on like the text around them (`Comment.anchor_text` reads through them, and a moved comment leaves no empty wrapper behind); a redline ignores a control that differs only by `w:id` or `w:tag` |
| the table of contents in a `w:sdt` with `w:docPartObj` | Word, Google Docs | one body block: `find_content_indices(heading)` returns it too |
| field instruction packed in one run (begin, instruction, separate), PAGE / NUMPAGES without a cached result | Google Docs | rdocx reads and refreshes them, and a redline after refreshing them compares |
| a default namespace on the root of each part | Google Docs | harmless for reading and most edits |
| measurements with a floating-point tail (`w:gridCol w:w="2210.0000000000005"`, `w:ind w:hanging="226.99999999999977"`, `w:trHeight`, `w:pgMar`) | Google Docs | rdocx reads each as the nearest integer |
| an empty `word/comments.xml` | Google Docs | kept byte for byte by a save |
| a `customXML` part with Google's round-trip data | Google Docs | keep it; rdocx does |
| several `w:style` elements with one id (`TableNormal`, `Normal`, `Table1`, ...), the later ones sometimes with other contents | Google Docs | `rebuild_toc()` and the style edits (`add_style()`) use the first definition |
| several default styles of one type under different ids (`TableNormal` and `TableauNormal`, up to four paragraph defaults) | Google Docs | `rebuild_toc()` and `add_style()` accept them |
| table of contents entries regenerated without their `TOC 1`, `TOC 2`... paragraph styles | Word, Google Docs, when they update the TOC | `rebuild_toc()` writes the entries back in the TOC styles |
| pictures re-encoded, `pageBreakBefore` and `keepNext` rewritten, styles added | Google Docs, on every save | nothing shows to the eye; see the next section |
| style ids in the interface language (`Titre1`, `Policepardfaut`) | Word in another language | assign styles by the id the file uses (`doc.styles`), never by an English name |
| `w:lineRule="auto"` spacing of 276 (Word) or 264 / 276 (Google) | both | rdocx lays lines at Word's heights; still check page breaks in the target application when they matter |

## What a save in Google Docs does

Observed behaviour of Google Docs on a .docx it opens; no test here can exercise it.

- **It re-serializes the whole package on every save**: content controls around comment anchors, fields packed,
  table-of-contents styles stripped, pictures re-encoded, `pageBreakBefore` and `keepNext` rewritten, styles
  added. Nothing shows to the eye. Resolving a thread in Google Docs saves again and drops the resolved comments
  from the .docx.
- **It keeps its own copy of the comments** of a .docx it has opened and writes it back at each save. A comment
  added with rdocx is taken in at Google's next read; a change to the author or the text of a comment Google
  already holds is undone at its next save, so such a change is made in Google Docs. The resolved state passes:
  a thread resolved with rdocx (`resolve_comment`) shows resolved in Google Docs, on a first open and on a file
  already open in Docs while the change syncs, and leaves the .docx at Google's next save, as above.
- **A read in Google Docs can change a header or footer unnoticed.** Before writing over a file someone has
  opened in Google Docs, compare its headers and footers with your last write: `rdocx diff LAST CURRENT` covers
  every story.

## Rules

- **Edit in place, never rebuild.** Opening, editing and saving with rdocx keeps every part you did not touch
  byte for byte, and everything it does not model inside the parts it rewrites, the root's `mc:Ignorable`
  included. Rebuilding a
  document from its text loses styles, numbering, fields, comments, bookmarks and producer data.
- **Do not round-trip through python-docx or LibreOffice** to finish an rdocx edit: each rewrites the whole
  package its own way. When a gap forces python-docx for one step, do that step on the file rdocx saved,
  and verify with rdocx afterwards.
- **A file annotated in Google Docs is rebased, never repaired.** Take the file as you last wrote it, before
  the reviewer's save; diff its text (`rdocx diff`) and its formatting (the runs of `rdocx text --json`, with
  their `formatting`) against the annotated file, to isolate the reviewer's own edits; make every edit on the
  clean copy; then carry the reviewer's comments across to it. Repairing the Google-saved file instead inherits
  everything its save rewrote (previous section).
- **Tracked changes and comments are shared state.** A file under review in Word or Google Docs carries the
  reviewers' threads and revisions: never accept, reject or remove them unless asked; add yours with an
  author name that says it is a machine (for example "Claude"), never the user's name, and a date (an undated
  comment shows none). `rdocx text`, `rdocx text --json` and `Paragraph.text` show the accepted view
  (insertions in, deletions out), and replacement edits the text inside an insertion, which stays tracked.
- **Fields and the TOC**: cached field results (TOC page numbers, PAGE, NUMPAGES) come from rdocx's
  pagination when you refresh them, which is close to Word's and not equal. When exact page numbers matter,
  say so and let Word update the fields (`doc.update_fields_on_open = True` asks Word to do it on open).
  Otherwise, after a save in Google Docs or Word, which leaves the page fields packed without a cached result,
  refresh the fields with `update_layout_backed_fields()` and set `doc.update_fields_on_open = False`
  (`<w:updateFields w:val="false"/>`): Word asks to update fields on open unless the settings say false
  (recipes.md, "Table of contents and page fields"). Word and Google Docs can also regenerate a table of
  contents without its TOC styles: `rebuild_toc()` writes them back.
- **Page count is a layout question first.** Before cutting text to save a page, list the active page breaks
  before paragraphs: Google Docs writes `pageBreakBefore` on every paragraph, mostly as `w:val="0"`, and
  `paragraph_format.page_break_before` is True only for an active one. Render a variant without each break and
  compare; a page count predicted without a render is often wrong (recipes.md, "Page breaks before cutting
  text").
- **Read text through rdocx, not the XML.** rdocx writes the apostrophe and the double quote as `&apos;` and
  `&quot;`: a check that searches `document.xml` as text does not find `owner's`.
- **Check in the target**: for a file going back to Word or Google Docs, the final check is opening it there;
  `rdocx validate` and a render catch most problems, not all.
