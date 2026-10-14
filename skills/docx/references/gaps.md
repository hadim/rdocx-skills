# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`paragraph.xml` / `replace_xml`, also on runs,
tables, cells and section properties: `references/python-api.md`, "Raw XML").

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A whole paragraph with a hyperlink [compare-hyperlink-paragraph] | `compare` (Python and CLI) refuses a pair in which a whole paragraph holding a hyperlink is inserted or deleted ("cannot track a whole paragraph's hyperlink or simple field"), at every granularity and with the ignore options; editing such a paragraph in place compares | in throwaway copies of both inputs, remove every hyperlink (`while doc.hyperlinks: doc.remove_hyperlink(doc.hyperlinks[0])`, the text stays), then compare: the redline has no links, say so |
| A paragraph next to a table [compare-paragraph-next-to-table] | `compare` refuses some pairs in which the paragraph right before a table is inserted or deleted ("needs an adjacent paragraph for a final paragraph change"): the first paragraph of the body, or a deletion next to an edited cell | add an empty paragraph before every table in both inputs (a blank line before each table in the redline) |
| An empty paragraph gains properties [compare-empty-paragraph-properties] | `compare` fails its own acceptance check ("acceptance does not reproduce the edited stories") when an empty paragraph gains paragraph properties (spacing, alignment), in the body or a header | `ignore_formatting=True` (`--ignore-formatting`), which leaves every formatting change untracked: say so |
| Paragraphs holding only a line break [compare-break-only-paragraphs] | `compare` fails its own acceptance check when two paragraphs holding only a line break (the answer boxes of a form), then a text paragraph, become one such paragraph | add an empty paragraph at the end of both inputs (one more blank paragraph in the redline) |
| Word changes interleaved [compare-word-grouping] | `granularity="word"` alternates one-word deletions and insertions across a rewritten stretch (`[-the-]{+to+} [-site-]{+the+}`), so a largely rewritten paragraph reads as word salad | for a document with heavy rewrites, `granularity="run"` replaces each changed run whole: readable, but a light edit replaces its whole run too |
| The tracked view in black only [tracked-view-colours] | `--revision-view tracked` and `revision_view="tracked"` draw insertions underlined and deletions struck through in black, with no colour option; raw XML refuses header and footer paragraphs (`NotImplementedError`), and `replace_xml` refuses a picture paragraph with an SVG twin whose `asvg` namespace is declared on the `w:p` | colour the runs inside `w:ins` (blue) and `w:del` (red) of a throwaway copy through `paragraph.xml` / `replace_xml`, then render it in the tracked view; revisions in headers and footers stay black, and keep the namespace declaration on `svgBlip` |

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).
