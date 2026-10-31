# bench: tools=rpptx render; processes=2; raw_xml=no
set -euo pipefail
rpptx render "$1" -o "$2" --dpi 100 --font-dir "$FONT_DIR"
