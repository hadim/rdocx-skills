#!/usr/bin/env bash
# bench: {"tools": ["accept_changes.py (LibreOffice)", "soffice.py (PDF)", "pdftoppm"], "processes": 6, "raw_xml": false, "libreoffice": true, "lo_profile": true, "notes": "LibreOffice's PDF shows tracked changes (coloured, underlined, struck through) by default: the accepted view needs accept_changes.py first"}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
mkdir -p "$2/accepted" "$2/tracked"
python "$S/accept_changes.py" "$1" "$W/accepted.docx" >/dev/null
cp "$1" "$W/tracked.docx"
python "$S/office/soffice.py" --headless --convert-to pdf --outdir "$W" "$W/accepted.docx" "$W/tracked.docx" >/dev/null 2>&1
pdftoppm -png -r 100 -f 1 -l 3 "$W/accepted.pdf" "$2/accepted/page"
pdftoppm -png -r 100 -f 1 -l 3 "$W/tracked.pdf" "$2/tracked/page"
