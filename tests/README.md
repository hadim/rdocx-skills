# Acceptance suite

```bash
python3 scripts/rdocx_env.py test                 # everything, about a minute and a half on linux-x86_64
python3 scripts/rdocx_env.py test -k pptx         # one vertical
python3 scripts/rdocx_env.py test --runxfail      # show how each known gap fails today
```

| File | Covers |
|---|---|
| `test_docx_read.py` | CLI text and JSON views, inspect, layout, structure, hyperlinks, conversion to md and html, stories other than the body, tracked changes in the views, validation, measurements with a decimal part |
| `test_docx_create.py` | a new document read back by python-docx, validated |
| `test_docx_edit.py` | counted replacement and its reach by story, structural edits, which calls invalidate handles, tables, pictures and their ids, styles, units, save semantics, byte stability |
| `test_docx_review.py` | comments (anchoring, dates, ids, table cells), threads, tracked changes, compare, the redline refusals |
| `test_docx_render.py` | fields, TOC, layout against Word's metrics, PDF and PNG, text layer, which outputs are overwritten or refused, CLI robustness |
| `test_docx_matrices.py` | identity attributes (18 rows) and producer traits (12 rows) by operation |
| `test_pptx.py` | reading, creating, editing, comments, notes, text fit, rendering, round trip |
| `test_skill_scripts.py` | the helpers shipped in `skills/*/scripts` |
| `test_installer.py` | `scripts/rdocx_env.py`: a fresh install, then each change `status` must catch and `install` repair |
| `test_release.py` | `scripts/rdocx_env.py` and the releases (upstream's and this repository's builds): download and check, a changed file refused, `lock --write --release` with its provenance check, finding the release tags of a commit, `bump` |
| `test_docs_snippets.py` | every block of `skills/*/references/recipes.md`, run in order, then checks on their outputs |
| `test_skill_docs.py` | the skills against the build: every cited Python name, keyword argument, CLI command and flag exists, the gap pages match `gaps.py`, each SKILL.md stays under 200 lines |

Fixtures are generated at run time by `fixtures/make_fixture_docx.py` and `fixtures/make_fixture_pptx.py`
(deterministic; invented text and drawn pictures), by `builders.py` (a document with a token in every story:
body, table cell, inline and block content controls, tracked insertion and deletion, text box, headers and
footers per variant, footnote, endnote; a paragraph with one wrapped run) and by small builders inside the
tests. Test dependencies
(pytest, python-docx, python-pptx, Pillow) are pinned by hash in `requirements.txt`; python-docx and
python-pptx serve only to build inputs and to re-read outputs independently.

**Gaps.** `@pytest.mark.gap("key")` turns a test into a strict expected failure, with its reason from
`gaps.py`. A gap that a new pin closes makes the run fail with `[XPASS(strict)]`: remove the marker, delete the
entry, update the skills. A gap about a missing API calls the name we expect: if the API lands under another
name, adapt the test.
