#!/usr/bin/env bash
# bench: {"tools": ["soffice.py (PDF)", "pdftotext", "unzip", "python + lxml (TOC by hand)", "zip"], "processes": 6, "raw_xml": true, "libreoffice": true, "files": ["builtin.sh", "builtin_edit.py"]}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
cp "$1" "$W/doc.docx"
python "$S/office/soffice.py" --headless --convert-to pdf --outdir "$W" "$W/doc.docx" >/dev/null 2>&1
pdftotext -layout "$W/doc.pdf" "$W/pages.txt"
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$(dirname "$0")/builtin_edit.py" "$W/u/word/document.xml" "$W/pages.txt"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
