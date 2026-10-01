#!/bin/sh
# Poppler (pdftotext, pdftoppm) for the PDF text and raster checks of the suite, which skip them without it: apt on
# Linux, Homebrew on macOS, MSYS2 on Windows (preinstalled on the runners). Its ucrt64/bin goes at the end of PATH,
# not first as GITHUB_PATH would put it: it also holds a MinGW python.exe, which must not replace the runner's.
set -eu
case "$RUNNER_OS" in
  Linux)
    sudo apt-get update -q > /dev/null
    sudo apt-get install -y -q --no-install-recommends poppler-utils > /dev/null ;;
  macOS)
    HOMEBREW_NO_AUTO_UPDATE=1 brew install --quiet poppler ;;
  Windows)
    /c/msys64/usr/bin/pacman -Sy --noconfirm --needed --noprogressbar mingw-w64-ucrt-x86_64-poppler > /dev/null
    python -c 'import os; print("PATH=" + os.environ["PATH"] + os.pathsep + r"C:\msys64\ucrt64\bin")' >> "$GITHUB_ENV" ;;
esac
