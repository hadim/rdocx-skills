#!/usr/bin/env bash
# Built-in skills' visual path: LibreOffice to PDF through the skill's soffice.py, then pdftoppm to PNG at 100 dpi.
#   lo_render.sh SKILL_DIR IN OUT_DIR [first]      (first: page or slide 1 only)
set -euo pipefail
W=$(mktemp -d)
python "$1/scripts/office/soffice.py" --headless --convert-to pdf --outdir "$W" "$2" >/dev/null 2>&1
mkdir -p "$3"
if [ "${4:-}" = "first" ]; then
  pdftoppm -png -r 100 -f 1 -l 1 "$W"/*.pdf "$3/page"
else
  pdftoppm -png -r 100 "$W"/*.pdf "$3/page"
fi
