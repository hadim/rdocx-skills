# bench: tools=rpptx slide import; processes=2; raw_xml=no
set -euo pipefail
rpptx slide import "$1" "$2" --slides 2-3 --at 2 --keep-source-formatting -o "$3"
