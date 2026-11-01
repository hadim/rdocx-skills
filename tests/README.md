# Acceptance suite

```bash
python3 plugin/scripts/rdocx_env.py test                                # everything
python3 plugin/scripts/rdocx_env.py test -n auto --dist loadfile        # in parallel, as CI does (a quarter faster)
python3 plugin/scripts/rdocx_env.py test -k pptx                        # one vertical
python3 plugin/scripts/rdocx_env.py test --runxfail                     # show how each known gap fails today
```

The suite covers the skills' own contract; rdocx's behaviour is upstream's to test. The behaviour files hold only
what the pinned build has and upstream `main` does not yet: a fix or a feature carried on an integration branch, or a
known gap (strict xfail). When the pin moves to an upstream release, run them against it and delete those that pass,
but `test_pptx_scaling.py` and `test_docx_scaling.py`, whose passing calls guard against a regression.

| File | Covers |
|---|---|
| `test_docs_snippets.py` | every block of `skills/*/references/recipes.md`, run in order, then checks on their outputs |
| `test_skill_docs.py` | the skills against the build: every cited Python name, keyword argument, CLI command and flag exists, every public class and CLI command is cited or excluded, the gap pages match `gaps.py`, each SKILL.md stays under 16 KB and its description under 800 characters |
| `test_skill_scripts.py` | the helpers shipped in `skills/*/scripts` |
| `test_installer.py` | `scripts/rdocx_env.py`: a fresh install, then each change `status` must catch and `install` repair, and the installs of older pins removed |
| `test_release.py` | `scripts/rdocx_env.py` and the releases (upstream's and this repository's builds): download and check, a changed file refused, `lock --write --release` with its provenance check, finding the release tags of a commit, `bump` |
| `test_ci_scripts.py` | `.github/scripts`: the suite summary's groups (gaps closed, docs not yet updated, other failures), the final suite written into the notes of this repository's release only, idempotent, the build-time suite of older releases kept |
| `test_docx_*.py`, `test_pptx*.py` | behaviour that upstream `main` does not have yet (see above) |
| `test_docx_scaling.py` | every public mutating call of rdocx, timed in a document of 25 paragraphs and in one of 100 (6 runs each with the run properties Google Docs writes, a picture and a table every 10th): no call grows with the document on the pin but the 15 of `ACCEPTED`, whose linear cost is accepted and bounded against a faster growth, so each one guards against a regression; also the reads `Paragraph.xml` and `ParagraphFormat.line_spacing`; `test_every_mutating_call_is_timed` checks the list against the stubs. A call over its bound is measured again, twice at most, the fastest of every measure counting, so that a load spike on a CI runner does not fail it. About 3.5 minutes on one worker of an M3 Max |
| `test_pptx_scaling.py` | every public mutating call of rpptx, timed in a deck of 5 slides and in one of 20: no call grows with the deck on the pin, so each one guards against a regression; `test_every_mutating_call_is_timed` checks the list against the stubs. A call over its bound is measured again, twice at most, the fastest of every measure counting, so that a load spike on a CI runner does not fail it. About 3.5 minutes on one worker of an M3 Max |

Fixtures are generated at run time by `fixtures/make_fixture_docx.py` and `fixtures/make_fixture_pptx.py`
(deterministic; invented text and drawn pictures), by `builders.py` (a document with a token in every story:
body, table cell, inline and block content controls, tracked insertion and deletion, text box, headers and
footers per variant, footnote, endnote; a paragraph with one wrapped run; decks of text boxes, tables and
pictures; documents of verbose paragraphs that end with the scaling test's targets) and by small builders inside
the tests. Test dependencies
(pytest, pytest-xdist, python-docx, python-pptx, Pillow) are pinned by hash in `requirements.txt`; python-docx and
python-pptx serve only to build inputs and to re-read outputs independently.

**Gaps.** `@pytest.mark.gap("key")` turns a test into a strict expected failure, with its reason from
`gaps.py`. A gap that a new pin closes makes the run fail with `[XPASS(strict)]`: remove the marker, delete the
entry, update the skills. A gap about a missing API calls the name we expect: if the API lands under another
name, adapt the test.
