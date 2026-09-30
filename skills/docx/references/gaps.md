# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Styles

| Gap | What happens | Workaround / fallback |
|---|---|---|
| Duplicate style ids [styles-duplicate-ids] | when `word/styles.xml` holds several `w:style` elements with one id (Google Docs writes `TableNormal`, `Normal`, `Table1` and others more than once), `add_style()` raises `RdocxError: invalid style graph: duplicate style ID '...'`; `rebuild_toc()` on the same file uses the first definition and says so in its `diagnostics` | before `add_style`, remove the later `w:style` elements of each repeated id with lxml, keeping the first, then go on with rdocx. No fallback needed beyond that one step |
