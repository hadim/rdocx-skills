# bench: tools=rpptx convert; processes=2; raw_xml=no
set -euo pipefail
rpptx convert "$1" --to md --notes -o "$2"
