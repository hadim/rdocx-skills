#!/usr/bin/env bash
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/unpacked"
find "$W/unpacked" -type l -delete
python "$BUILTIN_DOCX/scripts/merge_runs.py" "$W/unpacked"
python "$(dirname "$0")/builtin_fill.py" "$W/unpacked/word/document.xml" "$2"
(cd "$W/unpacked" && zip -qXr "$3" .)
