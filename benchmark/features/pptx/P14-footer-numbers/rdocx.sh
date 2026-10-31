# bench: tools=rpptx footer; processes=2; raw_xml=no
set -euo pipefail
rpptx footer "$1" --slide-number --footer "Northwind Confidential" --skip-title -o "$2"
