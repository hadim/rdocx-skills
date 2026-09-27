Comment run positions skip runs inside inline content controls, so comments anchor on the wrong text without an error

`RunPosition.run_index` (Python `add_comment`) and `rdocx comment add --start-run/--end-run` (CLI) count only
the paragraph's direct runs (`paragraph.runs` in `validate_run_range`, `crates/rdocx/src/comments.rs:1003`, and
the story variant at `:562`). The read APIs count the runs inside an inline `w:sdt` too: `Paragraph.runs` in
Python and `rdocx text --json` list them in document order. With an inline content control before the target,
a run index taken from either read API lands one run (or more) further right: the comment is written on the
wrong text, the command exits 0, and the last runs of the paragraph cannot be addressed at all. Inline content
controls are common in documents exported from Google Docs (`goog_rdk_*` controls) and in
templates.

Reproduction (python-docx builds the input; `rdocx` on PATH): `2026-09-27-comment-run-index-sdt.py`.

```
Paragraph.runs: ['before ', 'TARGET', ' after']
text --json runs: ['before ', 'TARGET', ' after']
python, runs [1, 2): ' after' expected 'TARGET'
cli, --start-run 1 --end-run 2: ' after' expected 'TARGET'
python, runs [2, 3): RdocxError comment range end run index 3 exceeds paragraph run count 2
```

The same happens with runs inside a tracked insertion: with the `w:sdt` wrapper of the reproduction replaced by
`<w:ins w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">` around the `TARGET` run, the output is identical
(the comment lands on `' after'`, run index 3 is refused). So a comment cannot be anchored reliably in any
paragraph under review.

*Acceptance:* one run index space for reading and anchoring: a run index read from `Paragraph.runs` or
`text --json` anchors on that run, including runs inside inline content controls and tracked insertions (the comment range markers
placed inside `w:sdtContent`, or around the whole `w:sdt` when the range covers it). If a range cannot be
anchored exactly, the call is refused with an error rather than shifted. A test per entry point (Python
`add_comment`, `add_story_comment`, CLI `comment add`) on a paragraph with an inline `w:sdt`, asserting the
anchored text.

Environment: `main` at `9a7ed714` (S75), release build, linux x86_64, python-docx 1.2.0 for the input.
