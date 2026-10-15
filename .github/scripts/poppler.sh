#!/bin/sh
# Poppler (pdftotext, pdftoppm) for the PDF text and raster checks of the suite, which skip them without it: apt on
# Linux, Homebrew on macOS, and on Windows a prebuilt Poppler pinned by SHA-256 (a few seconds, where MSYS2's pacman
# took about 3 minutes on the runners). Its bin folder goes at the end of PATH, not first as GITHUB_PATH would put it,
# so that nothing in it can shadow the runner's tools.
set -eu
command -v pdftotext > /dev/null && command -v pdftoppm > /dev/null && exit 0
case "$RUNNER_OS" in
  Linux)
    sudo apt-get update -q > /dev/null
    sudo apt-get install -y -q --no-install-recommends poppler-utils > /dev/null ;;
  macOS)
    HOMEBREW_NO_AUTO_UPDATE=1 brew install --quiet poppler ;;
  Windows)
    python - "$RUNNER_TEMP" >> "$GITHUB_ENV" <<'PY'
import hashlib, os, sys, urllib.request, zipfile
URL = ("https://github.com/oschwartz10612/poppler-windows/releases/download/v26.09.0-0/Release-26.09.0-0.zip")
SHA256 = "7a6f256a0ddf7536182246a5733331bf4677cbcc34f4663774947ad34556c8d0"
tmp = sys.argv[1]
archive = os.path.join(tmp, "poppler.zip")
urllib.request.urlretrieve(URL, archive)
with open(archive, "rb") as f:
    digest = hashlib.sha256(f.read()).hexdigest()
if digest != SHA256:
    sys.exit(f"{URL}: SHA-256 {digest}, expected {SHA256}")
zipfile.ZipFile(archive).extractall(os.path.join(tmp, "poppler"))
print("PATH=" + os.environ["PATH"] + os.pathsep + os.path.join(tmp, "poppler", "poppler-26.09.0", "Library", "bin"))
PY
    ;;
esac
