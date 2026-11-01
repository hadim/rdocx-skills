#!/usr/bin/env bash
# Build .cache/: the Python and node environments of the built-in and python-docx stacks,
# pandoc, and the metric-compatible fonts both renderers get. The rdocx stack is the pinned
# build: `python3 plugin/scripts/rdocx_env.py install` from the repository root.
set -euo pipefail
cd "$(dirname "$0")"
C=$PWD/.cache
mkdir -p "$C/bin" "$C/fonts" "$C/node"

# Python 3.13 with the libraries the built-in skills' scripts import, and python-docx / python-pptx.
if [ ! -x "$C/venv/bin/python" ]; then
  uv venv -q --python 3.13 "$C/venv"
fi
uv pip install -q --python "$C/venv/bin/python" \
  python-docx==1.2.0 python-pptx==1.0.2 lxml defusedxml Pillow numpy scikit-image openpyxl matplotlib

# docx-js and pptxgenjs, as the built-in skills use them.
(cd "$C/node" && [ -f package.json ] || npm init -y >/dev/null; npm install --silent docx@9.15.0 pptxgenjs@4.0.1 jszip)

# pandoc (the built-in docx skill reads with it).
if [ ! -x "$C/bin/pandoc" ]; then
  cp "$(pixi exec -s pandoc -- sh -c 'command -v pandoc')" "$C/bin/pandoc"
fi

# Carlito and Caladea (SIL OFL), metric-compatible with Calibri and Cambria.
for f in carlito/Carlito-Regular carlito/Carlito-Bold carlito/Carlito-Italic carlito/Carlito-BoldItalic \
         caladea/Caladea-Regular caladea/Caladea-Bold caladea/Caladea-Italic caladea/Caladea-BoldItalic; do
  [ -f "$C/fonts/$(basename $f).ttf" ] || curl -sSfL -o "$C/fonts/$(basename $f).ttf" \
    "https://github.com/google/fonts/raw/main/ofl/$f.ttf"
done
cat > "$C/fonts.conf" <<CONF
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>
  <dir>$C/fonts</dir>
  <cachedir>$C/fontcache</cachedir>
</fontconfig>
CONF
"$C/venv/bin/python" common.py
