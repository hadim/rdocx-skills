# Known gaps of the pinned build (pptx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so and this page is updated. "Fallback" means: do that step only with the default `pptx` skill
(python-pptx), and keep rpptx for the rest and for the verification.

| Gap | What happens | Workaround | Fallback |
|---|---|---|---|
| `Run.text` invalidates handles [pptx-run-text-stale] | after `run.text = ...`, every held slide, shape, paragraph and run handle raises `StaleElementError` | re-navigate from `prs` after each run edit | none needed |
| Inherited geometry [pptx-inherited-geometry] | a placeholder without its own `a:xfrm` (the normal case in a deck made in PowerPoint) reads `left`, `top`, `width`, `height` as None; setting one writes an `a:xfrm` with `y=0` and no size | read the layout placeholder's geometry (`slide.slide_layout` shapes) and set all four values together | python-pptx reads the inherited values |
| No slide duplication [pptx-duplicate-slide] | `duplicate_slide` exists in Rust, not in Python | add a slide on the same layout and rebuild its content, for simple slides | python-pptx has no API either: copy the slide XML with lxml |
| No replacement in Python [pptx-replace-python] | no `replace_text` | `pptx_ops.replace_batch` (the CLI with `--expect`, chained) | none needed |
| Comment resolve and remove in Python [pptx-comment-resolve-python] | Python has add, reply, move only | `rpptx comment resolve/remove` | none needed |
| Groups cannot be filled [pptx-group-population] | `add_group_shape()` returns an empty group whose `shapes` are read-only | add the shapes ungrouped | python-pptx `group_shape.shapes.add_*` |
| Table rows and columns [pptx-table-rows] | no row or column insertion or removal, no merge, fills or borders on cells from Python | size the table right when adding it | python-pptx for merges; lxml for rows |
| Z-order [pptx-zorder] | no bring forward or send backward; new shapes go on top | add shapes in drawing order | lxml: move the `p:sp` element in `p:spTree` |
| Built-in table styles [pptx-builtin-table-styles] | a table whose style is one of PowerPoint's built-in styles referenced by GUID, not defined in the file (python-pptx writes such tables), renders unstyled | set fills on the cells you need | render with LibreOffice when the look matters |
| PDF backgrounds [pptx-pdf-background] | solid (and gradient) slide backgrounds are drawn in PNG output and missing from the PDF: white text on a coloured slide disappears | check the PDF of any slide with a background; use PNG output when that is enough | LibreOffice for the PDF |
| Gradients without an angle [pptx-gradient-optional-attrs] | a gradient whose `a:lin` has no `ang` (python-pptx's `fill.gradient()` writes one) or whose `a:path` has no `path` makes the whole file refuse to open | none in rpptx | python-pptx for that deck, or add `ang="0"` with lxml first |
| Line pitch [pptx-line-pitch] | a paragraph with percentage line spacing (`a:lnSpc/a:spcPct`, python-pptx `line_spacing = 1.0`) is laid out at that percentage of the font size: 100 % gives 1.0 em where LibreOffice gives 1.2 em, less than the glyphs' own height. Lines overlap in the render, and `text_layout()` reports a `height` about 17 % short: a frame it says fits can overflow in another application | for frames with percentage spacing, compare `f.height * 1.2` with `f.usable.height` | render in LibreOffice for the final check of a tight frame |
| Line breaks [pptx-line-breaks] | in a paragraph whose direction is set (`rtl="0"`, which python-pptx's template sets in its text styles), lines break at any word boundary: before a comma, before a space (the next line starts with it) or before the hyphen of a compound. `text_layout()` and the render agree with each other, not with other applications | read `text_layout()` lines: one that starts with punctuation, a space or a hyphen is a wrong break; keep a width margin (`width_factor=0.95`) | none: say so when the break matters |
| Two `a:pPr` in a paragraph [pptx-duplicate-ppr] | a paragraph holding a second `a:pPr` (not schema-valid, met in real decks; python-pptx reads it) makes the whole file refuse to open: `XmlError ... DrawingML text contains duplicate pPr`, in Python and every CLI command | on a copy, remove every `a:pPr` of a paragraph after its first (lxml), then use rpptx on the copy | python-pptx for that deck |
| `convert` overwrites [cli-convert-overwrites] | `rpptx convert --to pdf` and `thumbnail` write over any existing output, the input included (image `convert` and `render` refuse) | always a new path | none needed |
| Hyperlinks [pptx-hyperlinks] | no hyperlink API on runs or shapes | none in rpptx | python-pptx `run.hyperlink.address` for that step |
| Template content type [template-save-as-document] | a .potx opened and saved as .pptx keeps the template content type (PowerPoint expects a presentation; not checked in PowerPoint) | `pptx_ops.fix_template_content_type(out)` after the save | none needed |
| Broken pipe [cli-broken-pipe] | `rpptx text F \| head` panics (exit 101) | write to a file | none needed |

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
