# bench: tools=soffice.py (LibreOffice), pdftoppm; processes=4; raw_xml=no
set -euo pipefail
tmp=$(mktemp -d)
python "$PPTX_SKILL/scripts/office/soffice.py" --headless --convert-to pdf --outdir "$tmp" "$1"
mkdir -p "$2"
pdftoppm -png -r 100 "$tmp"/*.pdf "$2/slide"
rm -rf "$tmp"
