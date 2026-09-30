# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Styles

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Duplicate style ids [styles-duplicate-ids] | when `word/styles.xml` holds several `w:style` elements with one id (Google Docs writes `TableNormal`, `Normal`, `Table1` and others more than once), `add_style()` raises `RdocxError: invalid style graph: duplicate style ID '...'`; `rebuild_toc()` on the same file uses the first definition and says so in its `diagnostics` | before `add_style`, repair the styles part with lxml as in the next row (keep the first `w:style` of each id), then go on with rdocx. No fallback needed beyond that one step |
| Several default styles of one type [styles-several-defaults] | when `word/styles.xml` marks several styles of one type as default under different ids (Google Docs writes a `TableNormal` and a localized `TableauNormal`, both `w:default="1"`, and up to four default paragraph styles), `add_style()` raises `RdocxError: invalid style graph: style type '...' has more than one default`, also once the duplicate ids are gone; `rebuild_toc()` accepts the file | one lxml pass over `word/styles.xml` before `add_style`: drop every `w:style` whose id was already seen, then remove `w:default` from every later default style of a type, keeping the first of each type. No fallback needed beyond that one step |
