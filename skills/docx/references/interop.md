# Files that travel between Word, Google Docs and LibreOffice

A .docx is rarely written by one application only. Each producer writes the same content its own way, and
the next one must still open what you save. What matters when you edit with rdocx:

## What each producer leaves in the XML

| Trait | Written by | What to do |
|---|---|---|
| `w:rsid*` on paragraphs, runs and rows, `w14:paraId` / `w14:textId`, a `w:rsids` list in settings | Word, on every save | carry no content; rdocx keeps them, except on table rows after an edit (gap tr-identity-lost). They break a plain TOC rebuild (gap toc-rsid-field-runs) |
| explicit `w:val="0"` toggles (`w:b`, `w:i`, `w:rtl`, `keepNext`, `pageBreakBefore`), `xml:space="preserve"` on every `w:t`, `w:orient="portrait"` | Google Docs | equivalent to the defaults; rdocx rewrites some of them in an edited `document.xml`, which shows as noise in a redline (gap compare-own-save-noise) |
| content controls `w:sdt` with `w:tag goog_rdk_N` around runs or paragraphs | Google Docs | text inside is invisible to replacements (gap sdt-replace) and, for paragraph-level controls, to plain `rdocx text` (gap sdt-text-cli); a comment after an inline control lands on the wrong run (gap comment-runposition-sdt: `docx_ops.comment_on_text` refuses); `add_picture` fails on such a file (gap add-picture-sdt-default-ns) |
| the table of contents in a `w:sdt` with `w:docPartObj` | Word, Google Docs | one body block: `find_content_indices(heading)` returns it too |
| field instruction packed in one run (begin, instruction, separate), PAGE / NUMPAGES without a cached result | Google Docs | rdocx reads and refreshes them; a redline after refreshing them is refused (gap compare-packed-fields) |
| a default namespace on the root of each part | Google Docs | harmless for reading and most edits |
| an empty `word/comments.xml` | Google Docs | re-serialised by any save (gap empty-comments-reserialised) |
| a `customXML` part with Google's round-trip data | Google Docs | keep it; rdocx does |
| style ids in the interface language (`Titre1`, `Policepardfaut`) | Word in another language | assign styles by the id the file uses (`doc.styles`), never by an English name |
| `w:lineRule="auto"` spacing of 276 (Word) or 264 / 276 (Google) | both | layout differs from Word's (gaps line-gap, picture-line-spacing): check page breaks in the target application when they matter |

## Rules

- **Edit in place, never rebuild.** Opening, editing and saving with rdocx keeps every part you did not touch
  byte for byte, and everything it does not model inside the parts it rewrites. Rebuilding a document from
  its text loses styles, numbering, fields, comments, bookmarks and producer data.
- **Do not round-trip through python-docx or LibreOffice** to finish an rdocx edit: each rewrites the whole
  package its own way. When a gap forces python-docx for one step, do that step on the file rdocx saved,
  and verify with rdocx afterwards.
- **Tracked changes and comments are shared state.** A file under review in Word or Google Docs carries the
  reviewers' threads and revisions: never accept, reject or remove them unless asked; add yours with an
  author name that says it is a machine (for example "Claude"), never the user's name, and a date (an undated
  comment shows none). Read such a file with `rdocx text --json` or
  `Paragraph.text` (accepted view): plain `rdocx text` drops inserted text, and replacements do not reach it.
- **Fields and the TOC**: cached field results (TOC page numbers, PAGE, NUMPAGES) come from rdocx's
  pagination when you refresh them, which is close to Word's and not equal. When exact page numbers matter,
  say so and let Word update the fields (`doc.update_fields_on_open = True` asks Word to do it on open).
- **Check in the target**: for a file going back to Word or Google Docs, the final check is opening it there;
  `rdocx validate` and a render catch most problems, not all.
