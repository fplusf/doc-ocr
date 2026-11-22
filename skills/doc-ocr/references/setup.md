# doc-ocr setup

Run `bash scripts/setup.sh` to see what is available. It changes nothing.

## Required
- Python 3.9+
- PyMuPDF: `pip3 install pymupdf` (PDF reading, page rendering, all image formats)
- python-docx: `pip3 install python-docx` (Word input and export)

## Optional
- openpyxl (Excel input), beautifulsoup4 (HTML/EPUB input)
- Pillow is not needed. If present it is only used to rotate JPG photos on non-Mac systems.

## OCR engines
| Engine | Where | Good at | Languages |
|---|---|---|---|
| Apple Vision | macOS 13+ (compiled from `scripts/vision_ocr.swift` on first use, needs `swiftc`) | print, photos, screenshots, light handwriting | en, ru, uk, de, fr, es, it, pt, nl, pl, cs, ro, sv, da, no, tr, id, ms, vi, th, ar, zh, ja, ko |
| Tesseract | any OS: `brew install tesseract` / `apt install tesseract-ocr` | clean scans, any language with a pack | 100+ with packs |
| Claude vision | always | handwriting, stamps, tables, damaged pages | any |

## Tesseract language packs
Vision has no Tajik, Uzbek, Kazakh, Kyrgyz, Persian or Hindi model. For those, add Tesseract packs (codes: `tgk`, `uzb`, `uzb_cyrl`, `kaz`, `kir`, `fas`, `hin`, `rus`, `eng`):

```bash
bash scripts/setup.sh --langs tgk rus
```

This downloads from github.com/tesseract-ocr/tessdata_best into the tessdata folder. Homebrew users can instead run `brew install tesseract-lang` (all languages, about 650 MB).

Short code to Tesseract code mapping lives in `LANGS` at the top of `scripts/extract.py`; you can also pass a raw Tesseract code (e.g. `--lang uzb_cyrl`).
