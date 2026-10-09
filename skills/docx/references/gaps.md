# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Rendering

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Fonts from a folder in images [render-font-dir-docx] | `--font-dir` / `font_dir=` reach the PDF only. `rdocx convert --to png/jpeg/tiff --font-dir` ignores the folder without a word, `rdocx render` has no `--font-dir`, and `render_page_to_png`, `render_all_pages`, `render_pages` take no fonts: a font that is not installed falls back to a bundled sans-serif in every image | check the look on the PDF made with `--font-dir`; for PNG previews, rasterise that PDF (`pdftoppm -png -r 150`), or install the fonts for the user first, then render |
| Bold from a variable font [variable-font-bold] | a variable font in the `--font-dir` folder (Google Fonts ships most families as `Family[wght].ttf`) gives a bold run the regular glyphs, under a `-Bold` font name in the PDF | put the static `Family-Bold.ttf` (and `-Italic`, `-BoldItalic`) in the folder when the family has them, else say that bold shows as regular |

## Validation

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A paragraph holding only a picture [validate-picture-paragraph] | `rdocx validate` counts it in "N empty paragraph(s) found" (a warning, exit 0) | subtract the paragraphs that hold a picture before acting on that warning; never remove them to silence it |

## Python API

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Colour arguments [colour-forms-docx] | `add_style` / `set_style(color=)` and `font.color` take an `rdocx.RGBColor` only (a hex string raises `'str' object is not an instance of 'tuple'`); `set_borders`, `set_border` and `shading` take a hex string only (`'RGBColor' object is not an instance of 'str'`); python-docx's `font.color.rgb = ...` raises `AttributeError` | `RGBColor(r, g, b)` (or `RGBColor.from_string("7B1E3A")`) for styles and fonts, `"7B1E3A"` for borders and shading; set `font.color = RGBColor(...)` directly |
| Table indent [table-indent] | `Table` has no `indent` (`w:tblInd`): a table whose first cell has a left margin stands out of the text column by that margin in a document below compatibility mode 15, as in Word | add `<w:tblInd w:w="N" w:type="dxa"/>` to the table's `w:tblPr` with python-docx or lxml on the saved file (`N` in twips, negative to pull the table left), then reopen with rdocx; indenting the cells' paragraphs moves the text, not the fill |

## Editing

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Replacement in one text box [replace-text-at-text-box-copy] | `replace_text_at` on a text-box paragraph edits Word's copy of the box (`mc:Choice`) and leaves the VML fallback copy (`mc:Fallback`) with the old text | use the document-wide `try_replace_text` / `replace_text` with `expect=`, which edits both copies and counts the box once |
