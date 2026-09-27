#!/bin/sh
# shellcheck disable=SC2016  # the backquotes are Markdown, not command substitutions
# Markdown summary of an acceptance suite log: the final count, gaps closed (strict xfails that now pass) and
# other failures. Usage: suite-summary.sh PYTEST_LOG PLATFORM
log="$1"
echo "### $2"
echo
if [ ! -s "$log" ]; then
  echo "The suite did not run."
  exit 0
fi
echo "\`$(tail -n 1 "$log" | tr -d '=' | sed 's/^ *//; s/ *$//')\`"
closed=$(grep -F '[XPASS(strict)]' "$log" | sed 's/^FAILED //' || true)
other=$(grep -E '^(FAILED|ERROR) ' "$log" | grep -vF '[XPASS(strict)]' || true)
if [ -n "$closed" ]; then
  printf '\nGaps closed (remove the marker and the entry in tests/gaps.py, then update the skills):\n\n```\n%s\n```\n' "$closed"
fi
if [ -n "$other" ]; then
  printf '\nOther failures (regressions or changed APIs):\n\n```\n%s\n```\n' "$other"
fi
