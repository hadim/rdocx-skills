# Findings

Gaps met while using rdocx on real work, before they are triaged. One file per finding,
`YYYY-MM-DD-<slug>.md`, written as the upstream ticket would be: what fails, a neutral reproduction that
builds its own input and prints the wrong result next to the expected one, the output of a real run, an
acceptance criterion, the pinned commit. Never the content, names or paths of a real document.

The skills ask agents to leave findings as `rdocx-findings/` in the folder they work in; copy them here.

Triage, for each finding:

1. Check `tests/gaps.py` and the upstream tracker for a duplicate.
2. Add the test: a strict xfail with `@pytest.mark.gap("key")` and an entry in `tests/gaps.py`.
3. Add the gap to the skill's `references/gaps.md`, with its workaround and fallback.
4. File the ticket upstream, put its number in `tests/gaps.py`, and delete the finding file.

| Finding | Status |
|---|---|
| `2026-09-27-convert-overwrites.md` | test and gap entry added; filed 27/09, number to add |
| `2026-09-27-add-picture-default-namespace.md` | test and gap entry added; filed 27/09, number to add |
| `2026-09-27-comment-run-index-sdt.md` | test and gap entry added; filed 27/09, number to add |
