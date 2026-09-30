# Findings

Gaps met while using rdocx on real work, before they are triaged. One file per finding,
`YYYY-MM-DD-<slug>.md`: what fails, a neutral reproduction that builds its own input and prints the wrong
result next to the expected one, the output of a real run, an acceptance criterion, the pinned commit. Never
the content, names or paths of a real document.

The skills ask agents to leave findings as `rdocx-findings/` in the folder they work in; copy them here.

Triage, for each finding:

1. Check `tests/gaps.py` for a duplicate.
2. Add the test: a strict xfail with `@pytest.mark.gap("key")` and an entry in `tests/gaps.py`.
3. Add the gap to the skill's `references/gaps.md`, with its workaround and fallback.
4. Delete the finding file: from then on the test is the reproduction, and the suite run on each new build
   says when the gap closes.

A finding is written so that it can be filed upstream as it is. Once filed, the issue is named in a comment on
the gap's entry in `tests/gaps.py` (step 2), not in the finding.
