#!/bin/sh
# shellcheck disable=SC2016  # the backquotes are Markdown, not command substitutions
# Markdown summary of an acceptance suite log: the final count, gaps closed (strict xfails that now pass), failures
# of tests/test_skill_docs.py (the skills not yet documenting a new build) and other failures.
# Usage: suite-summary.sh PYTEST_LOG HEADING (the platform, and how its files were installed when it matters)
log="$1"
echo "### $2"
echo
if [ ! -s "$log" ]; then
  echo "The suite did not run."
  exit 0
fi
echo "\`$(tail -n 1 "$log" | tr -d '=' | sed 's/^ *//; s/ *$//')\`"
closed=$(grep -F '[XPASS(strict)]' "$log" | sed 's/^FAILED //' || true)
failed=$(grep -E '^(FAILED|ERROR) ' "$log" | grep -vF '[XPASS(strict)]' || true)
docs=$(printf '%s\n' "$failed" | grep -E '^(FAILED|ERROR) tests/test_skill_docs\.py' || true)
other=$(printf '%s\n' "$failed" | grep -vE '^(FAILED|ERROR) tests/test_skill_docs\.py' | grep . || true)
if [ -n "$closed" ]; then
  printf '\nGaps closed (remove the marker, and the entry in tests/gaps.py once no test carries it, then update the skills):\n\n```\n%s\n```\n' "$closed"
fi
if [ -n "$docs" ]; then
  printf '\nDocs not yet updated for the new build (tests/test_skill_docs.py):\n\n```\n%s\n```\n' "$docs"
fi
if [ -n "$other" ]; then
  printf '\nOther failures (regressions or changed APIs):\n\n```\n%s\n```\n' "$other"
fi
