# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Comparison and rendering

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A table added or removed at the end of the body [compare-final-table] | when the edited version adds (or removes) a table followed by a paragraph at the very end of the body, `compare()` and `rdocx compare` refuse the pair: `comparison needs an adjacent paragraph for a final paragraph change`. The same table between two paragraphs compares | on copies of both versions, never on the inputs: add one empty paragraph at the end of each (`doc.add_paragraph("")`), then compare the copies. No fallback needed beyond that one step |
| A picture whose image changed [compare-picture-change] | when a figure keeps its place but gets a new image, `compare()` records no revision and the redline keeps the old image if nothing else changed; if anything else changed too, it refuses the pair: `comparison acceptance does not reproduce the edited stories at body story item[N]` | before comparing, check whether the pictures differ (`image_data(rid)` of each `drawing` item on both sides). If one does, put the new image into a copy of the original (`replace_image(rid, bytes)`), compare that copy, and anchor a comment on the caption saying the figure changed: the redline cannot show it as a revision |
| A PDF that shows the tracked changes [render-tracked-view] | `to_pdf()`, `render_pages()`, `rdocx convert --to pdf` and `rdocx render` render the accepted view only: insertions in, deletions out, no marks. The Rust API has a tracked view; Python and the CLI do not expose it | on a copy of the redline, with lxml, give the runs inside `w:ins` an underline and a colour and turn each `w:del` into a struck-through, coloured run (`w:delText` becomes `w:t`), then render the copy with rdocx. Say in the answer that the PDF is a rendering of marks, not Word's own |

