#!/usr/bin/env bash
# bench: {"tools": ["docx_ops.py comment"], "processes": 1, "raw_xml": false}
set -euo pipefail
python -m docx_ops comment "$1" "$2" --anchor "thirty days after delivery" \
  --text "Should this be 45 days?" --author "Dana Reviewer" --initials DR --date 2026-10-30T09:00:00Z
