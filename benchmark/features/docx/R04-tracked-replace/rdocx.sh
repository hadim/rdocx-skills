#!/usr/bin/env bash
# bench: {"tools": ["rdocx replace", "rdocx compare --granularity word"], "processes": 2, "raw_xml": false}
set -euo pipefail
T=$(mktemp -d)
rdocx replace "$1" -p thirty -v forty-five --expect 1 -o "$T/edited.docx"
rdocx compare "$1" "$T/edited.docx" --author "Dana Reviewer" --timestamp 2026-10-30T11:00:00Z \
  --granularity word -o "$2" >/dev/null
