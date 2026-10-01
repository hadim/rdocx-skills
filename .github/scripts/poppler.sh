#!/bin/sh
# Poppler (pdftotext, pdftoppm) for the PDF text and raster checks of the suite, which skip them without it: apt on
# Linux, Homebrew on macOS, MSYS2 on Windows (preinstalled on the runners; its ucrt64/bin goes on PATH).
set -eu
case "$RUNNER_OS" in
  Linux)
    sudo apt-get update -q > /dev/null
    sudo apt-get install -y -q --no-install-recommends poppler-utils > /dev/null ;;
  macOS)
    HOMEBREW_NO_AUTO_UPDATE=1 brew install --quiet poppler ;;
  Windows)
    /c/msys64/usr/bin/pacman -Sy --noconfirm --needed --noprogressbar mingw-w64-ucrt-x86_64-poppler > /dev/null
    cygpath -w /c/msys64/ucrt64/bin >> "$GITHUB_PATH" ;;
esac
