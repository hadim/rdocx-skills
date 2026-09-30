# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Opening a file

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Measurements with a decimal part [decimal-measurements] | Google Docs writes measurements with a floating-point tail (`w:gridCol w:w="2210.0000000000005"`, `w:ind w:hanging="226.99999999999977"`, `w:trHeight`, `w:pgMar`); such a file does not open: `Document.open()` and every CLI command, `rdocx text` and `validate` included, fail with `OXML parsing error: parse int error: invalid digit found in string` (no attribute named) or `unsupported table measurement: "..."`. Even `240.0` is refused, in the body, the styles and the numbering; only `w:spacing/@w:line` and whole table widths (`w:tcW w:w="4320.0"`) open | on a copy, never on the input: in the `word/*.xml` parts, round every `w:` attribute whose value is a decimal number (`-?\d+\.\d+`) to the nearest integer (zipfile and one regular expression), then open the copy with rdocx and go on. No fallback needed beyond that one step |
| Two drawings of one part with one id [docpr-duplicate-ids] | a file in which two drawings of one part share a `wp:docPr` id (several `id="0"` in the body) does not open: `cannot scan identifiers in XML part /word/document.xml: duplicate drawing id 0 in imported or preserved XML`, in `Document.open()` and every CLI command. One id used once in the body and once in a header opens | on a copy: in each part, give every `wp:docPr` whose id was already used earlier in that part a new id above the largest one of the package, then open the copy with rdocx. No fallback needed beyond that one step |

## Styles

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Duplicate style ids [styles-duplicate-ids] | when `word/styles.xml` holds several `w:style` elements with one id (Google Docs writes `TableNormal`, `Normal`, `Table1` and others more than once), `add_style()` raises `RdocxError: invalid style graph: duplicate style ID '...'`; `rebuild_toc()` on the same file uses the first definition and says so in its `diagnostics` | before `add_style`, repair the styles part with lxml as in the next row (keep the first `w:style` of each id), then go on with rdocx. No fallback needed beyond that one step |
| Several default styles of one type [styles-several-defaults] | when `word/styles.xml` marks several styles of one type as default under different ids (Google Docs writes a `TableNormal` and a localized `TableauNormal`, both `w:default="1"`, and up to four default paragraph styles), `add_style()` raises `RdocxError: invalid style graph: style type '...' has more than one default`, also once the duplicate ids are gone; `rebuild_toc()` accepts the file | one lxml pass over `word/styles.xml` before `add_style`: drop every `w:style` whose id was already seen, then remove `w:default` from every later default style of a type, keeping the first of each type. No fallback needed beyond that one step |

## Saving

| Gap | What happens | Workaround / fallback |
|---|---|---|
| An edit re-serializes its part [edit-reserializes-part] | every other part keeps its bytes and nothing is lost, but the part an edit touches (`word/document.xml` for the body) is written again in full: indented, with `xmlns:w` declared again on every element that carries `w:rsid*` (Word writes them on nearly every paragraph and run). The part grows, and an XML diff shows all of it instead of the edit | review an edit by text (`rdocx diff`, `rdocx text --json`), never by the XML of the parts. No fallback needed: the content is kept |
