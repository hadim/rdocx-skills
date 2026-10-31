#!/usr/bin/env bash
# bench: {"tools": ["rdocx convert --to pdf", "rdocx convert --revision-view tracked"], "processes": 2, "raw_xml": false}
set -euo pipefail
mkdir -p "$2"
rdocx convert "$1" --to pdf -o "$2/document.pdf" --font-dir "$FONT_DIR" >/dev/null
rdocx convert "$1" --to pdf --revision-view tracked -o "$2/tracked.pdf" --font-dir "$FONT_DIR" >/dev/null
