#!/usr/bin/env bash
# bench: {"tools": ["soffice.py (PDF)", "pdftotext", "pandoc -t json", "python"], "processes": 5, "raw_xml": false, "libreoffice": true, "files": ["builtin.sh", "builtin_pages.py"]}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
cp "$1" "$W/doc.docx"
python "$S/office/soffice.py" --headless --convert-to pdf --outdir "$W" "$W/doc.docx" >/dev/null 2>&1
pdftotext -layout "$W/doc.pdf" "$W/pages.txt"
pandoc -t json "$1" -o "$W/doc.json"
python "$(dirname "$0")/builtin_pages.py" "$W/doc.json" "$W/pages.txt" "$2"
