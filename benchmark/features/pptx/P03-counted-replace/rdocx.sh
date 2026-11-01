# bench: tools=rpptx replace; processes=2; raw_xml=no
set -euo pipefail
rpptx replace "$1" -p "$3" -v "$4" --expect "$5" -o "$2"
