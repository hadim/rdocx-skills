#!/usr/bin/env bash
# bench: {"tools": ["rdocx compare --granularity word"], "processes": 1, "raw_xml": false, "notes": "accepted and rejected views are right, but words that v2 merely splits into other runs (an unchanged paragraph included) come out deleted and inserted again (NOTES_review.md, finding 2)"}
set -euo pipefail
rdocx compare "$1" "$2" --author "Dana Reviewer" --timestamp 2026-10-30T12:00:00Z --granularity word -o "$3" >/dev/null
