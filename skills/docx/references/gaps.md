# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

## Redline

| Gap | What happens | Workaround / fallback |
|---|---|---|
| TOC rebuilt before the redline [compare-rebuilt-toc] | `compare()` and `rdocx compare` refuse the pair when the edited side had its TOC rebuilt | compare before rebuilding the TOC. If it still refuses, report the redline as unavailable. Fallback for a human-facing redline: Word's Compare, done by the user |

## Fields, layout, rendering

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Line gap [line-gap] | a single line is laid out without the font's line gap: Calibri 1.000 em where Word uses 1.221 em, Arial 1.117 against 1.150 | page counts and page breaks differ from Word, most with Calibri: never state a Word page number from rdocx's layout; check in Word when it matters |
| Picture lines [picture-line-spacing] | a line holding an inline picture is multiplied by proportional spacing (400 pt picture at line 264: 440 pt; Word 402.7 pt) | tall figures can be split from their caption in rdocx's render only |
| Tab stops [tab-stops] | text after a tab stop set in the paragraph or its style starts 36 pt (half an inch) before the stop, and a right-aligned stop is laid out as a left one: a page number after a dot leader ends short of its stop by 36 pt minus its own width | rdocx's render only; the file is untouched. Never judge tab alignment (TOC, forms, signature lines) from rdocx's PDF or PNG; check it in the target application |
| Numbered TOC entries [toc-numbered-entries] | for a numbered heading, `rebuild_toc()` writes number, tab, title, tab, page and adds no stop for the first tab: when the TOC style has only its right page-number stop, the title is pushed to the right margin | before rebuilding, give the TOC styles a left tab stop after the number (python-docx, one step on the styles), or leave the TOC to Word (`doc.update_fields_on_open = True`); rdocx's own render of the entries stays off (gap tab-stops) |
