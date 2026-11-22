#!/usr/bin/env bash
# Check what doc-ocr can use on this machine, and optionally add Tesseract language packs.
#   setup.sh                 report only (changes nothing)
#   setup.sh --langs tgk rus  download those Tesseract packs (tessdata_best) into the tessdata folder
set -u
here="$(cd "$(dirname "$0")" && pwd)"
ok()   { printf '  [ok]      %s\n' "$1"; }
miss() { printf '  [missing] %s\n' "$1"; }

echo "doc-ocr environment check"
# A universal (arm64 + x86_64) Python can load packages built for only one CPU type.
# Pick the launch command under which PyMuPDF actually imports.
PY=python3
for cand in "python3" "arch -arm64 python3" "arch -x86_64 python3"; do
  if $cand -c 'import fitz' 2>/dev/null; then PY="$cand"; break; fi
done
command -v python3 >/dev/null && ok "python3 $(python3 -V 2>&1 | cut -d' ' -f2), run scripts with: $PY" || miss "python3"
$PY -c 'import fitz' 2>/dev/null && ok "PyMuPDF (PDF + image handling)" || miss "PyMuPDF: pip3 install pymupdf"
$PY -c 'import docx' 2>/dev/null && ok "python-docx (Word export)" || miss "python-docx: pip3 install python-docx"
$PY -c 'import openpyxl' 2>/dev/null && ok "openpyxl (Excel input)" || miss "openpyxl (only for .xlsx input): pip3 install openpyxl"
$PY -c 'import bs4' 2>/dev/null && ok "beautifulsoup4 (HTML/EPUB input)" || miss "beautifulsoup4 (only for HTML/EPUB): pip3 install beautifulsoup4"
if [ "$(uname)" = "Darwin" ]; then
  if command -v swiftc >/dev/null; then ok "Apple Vision OCR (compiled on first use)"; else miss "swiftc for Apple Vision: xcode-select --install"; fi
fi
if command -v tesseract >/dev/null; then
  ok "tesseract $(tesseract --version 2>&1 | head -1 | cut -d' ' -f2)"
  echo "            languages: $(tesseract --list-langs 2>/dev/null | tail -n +2 | tr '\n' ' ')"
else
  miss "tesseract (needed for Tajik, Uzbek, Kazakh, Persian, Hindi...): brew install tesseract"
fi

if [ "${1:-}" = "--langs" ]; then
  shift
  command -v tesseract >/dev/null || { echo "install tesseract first"; exit 1; }
  dir="${TESSDATA_PREFIX:-$(tesseract --list-langs 2>/dev/null | head -1 | sed -E 's/.*"(.*)".*/\1/')}"
  [ -d "$dir" ] || { echo "cannot find tessdata folder ($dir)"; exit 1; }
  for l in "$@"; do
    echo "downloading $l.traineddata into $dir"
    curl -fL --retry 2 -o "$dir/$l.traineddata" "https://github.com/tesseract-ocr/tessdata_best/raw/main/$l.traineddata" \
      && ok "$l" || miss "$l (download failed)"
  done
fi
