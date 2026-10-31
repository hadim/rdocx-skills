#!/usr/bin/env bash
# bench: {"tools": ["rdocx render", "rdocx render --revision-view tracked"], "processes": 2, "raw_xml": false}
set -euo pipefail
rdocx render "$1" -o "$2/accepted" --pages 1-3 --dpi 100 --font-dir "$FONT_DIR" >/dev/null
rdocx render "$1" -o "$2/tracked" --pages 1-3 --dpi 100 --revision-view tracked --font-dir "$FONT_DIR" >/dev/null
