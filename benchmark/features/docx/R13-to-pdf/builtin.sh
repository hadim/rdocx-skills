#!/usr/bin/env bash
# bench: {"tools": ["accept_changes.py (LibreOffice)", "soffice.py (PDF)"], "processes": 4, "raw_xml": false, "libreoffice": true, "lo_profile": true, "notes": "LibreOffice's PDF shows tracked changes by default: the accepted view needs accept_changes.py first"}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
mkdir -p "$2"
python "$S/accept_changes.py" "$1" "$W/document.docx" >/dev/null
cp "$1" "$W/tracked.docx"
python "$S/office/soffice.py" --headless --convert-to pdf --outdir "$2" "$W/document.docx" "$W/tracked.docx" >/dev/null 2>&1
