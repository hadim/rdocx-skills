#!/usr/bin/env bash
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/unpacked"
find "$W/unpacked" -type l -delete
python "$BUILTIN_DOCX/scripts/merge_runs.py" "$W/unpacked"
python "$(dirname "$0")/builtin_replace.py" "$W/unpacked" "Service Level Agreement" "Support Contract" "$3"
(cd "$W/unpacked" && zip -qXr "$2" .)
