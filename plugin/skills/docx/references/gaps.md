# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and `doc.section(i)`: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| python-docx's names [python-docx-names] | most python-docx code runs, but these fail with a bare `AttributeError`, `TypeError` or `ImportError`: `style="List Bullet"` or `"List Number"` on a new document (`KeyError`), a margin set on `doc.sections[i]` (`left_margin`...), `MSO_COLOR_TYPE` for `font.color.type`, a settable `TabStop`, `custom_properties[k]`, `add_picture` with a width alone, `paragraph._p` | write it the rdocx way (`references/python-api.md`, "python-docx names"): `add_bullet_list_item` / `add_numbered_list_item`, `doc.section(i).left_margin = x`, `font.color.type == "rgb"`, `del tab_stops[i]` then `add_tab_stop`, `custom_properties.get/set/delete`, both picture sizes from the image's aspect ratio (`docx_ops.image_size`), `paragraph.xml` / `replace_xml` |
| A bare int length saves as 0 [docx-bare-int-length-saved-as-zero] | a bare integer is read as EMU: `font.size = 12` raises, but `space_after = 6`, `character_spacing = 2`, `update_section(0, margin_top=1440)` and `line_spacing = 2` are accepted and saved as about 0, with no error | always build lengths with `rdocx.Pt`, `Inches`, `Cm`, `Mm` or `Twips`, and a line spacing multiple as a float (`2.0`); read the value back after setting it |
| `validate` refuses Word's table of contents [docx-validate-paraid-in-content-control] | `rdocx validate` and `doc.validate()` fail with "main document paragraph owners do not match validation scope" on a document holding a paragraph with `w14:paraId` or `w14:textId` inside a block content control, which is how Word writes every table of contents and Google Docs its wrappers | `docx_ops.py validate F` validates a scratch copy without those two attributes (the snippet below does the same), only for the check: deliver the original. Also open the file with python-docx and render it |
| Notes: reference style, removal [docx-note-reference-and-removal] | `add_footnote` / `add_endnote` write the reference without Word's `FootnoteReference` / `EndnoteReference` style, so Word shows the number full size; a note handle retires at the next structural edit and nothing lists the notes, so a note cannot be removed later | right after adding a note, raise its mark: `p.runs[k].font.superscript = True` on the run that holds the reference (its `text` is `""`, right after the anchor); delete a note with `note.delete()` on the handle `add_footnote` returned, before any other edit; to remove a note of an existing document, fallback |
| Content control values [docx-content-control-value-like-word] | `set_content_control_value_by_tag` / `_by_alias` (and `rdocx fill`) keep `w:showingPlcHdr`, so Word shows the value as placeholder text, write no `w:fullDate` for a date, write a drop-down value as given (not its display text, not checked against the items), and return 0 for an unknown tag | check the returned count; pass a drop-down's display text, one of its items; then edit the control's XML in its paragraph: take `p.xml`, drop `<w:showingPlcHdr/>` and, for a date, add `w:fullDate="2026-10-31T00:00:00Z"` to its `w:date` element, and write it back with `p.replace_xml(xml)` |
| Restarting a list, a bullet's font [docx-list-numbering-options] | no `restart=` on `add_numbered_list_item`, no `restart_numbering`, no `font=` on `ListLevel` | a new count: `n = doc.add_numbering_instance(doc.add_numbering_definition([rdocx.ListLevel(format="decimal", text="%1.")]), start=1)`, then `p.numbering = (n, 0)` on its items (its indents are the level's: set `left_indent=`, `hanging_indent=`); a bullet font: fallback on `word/numbering.xml` |
| Raw XML that Word refuses [docx-raw-xml-unchecked] | `replace_xml` accepts an unknown `w:` element and a foreign element whose namespace the root's `mc:Ignorable` does not list; Word then reports the file as damaged, and `rdocx validate` does not flag it | write only elements of the WordprocessingML schema, in schema order, and foreign markup only inside `mc:AlternateContent` or in a namespace `mc:Ignorable` lists; open the result in Word or LibreOffice when it matters |
| Reading charts back [docx-chart-read-back] | a Word document's charts are written (`add_chart`) but not read: no `doc.charts` | read the chart part with `zipfile`: `word/charts/chart1.xml` for the first chart (the part the `r:id` of the `c:chart` in the paragraph's `xml` names) holds the categories and values (`c:cat`, `c:val`, `c:v`); the legend reads with `doc.chart_legend(paragraph)` |
| Display equations [docx-display-equation] | `add_equation` inserts inline equations only (`is_display` False) | insert it inline in an empty paragraph, then wrap its `m:oMath` in `m:oMathPara` with `p.replace_xml(...)` (declare `xmlns:m` on the new element); `equations()[0].is_display` then reads True |
| Values accepted silently [docx-invalid-values-accepted] | no error for `doc.save("x.pdf")` (a DOCX under that name), a second tab stop at one position, a page break in a header or footer (Word ignores it), a malformed language tag (`font.language = "not a tag!"`), a font face without its `name` table | save as `.docx` and convert (`to_pdf()`, `rdocx convert`); check a paragraph's `tab_stops` before adding one; use a BCP 47 tag such as `"fr-FR"`; embed only a whole TrueType or OpenType face and check `doc.fonts` afterwards |

The scratch copy for `rdocx validate` ([docx-validate-paraid-in-content-control] above):

```python
import re, zipfile
with zipfile.ZipFile("out.docx") as src, zipfile.ZipFile("check.docx", "w") as dst:
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename == "word/document.xml":
            data = re.sub(rb' w14:(?:paraId|textId)="[^"]*"', b"", data)
        dst.writestr(info, data)
```

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
