# doc-ocr

A Claude skill that reads text out of any document, including scans and pictures you cannot copy from, and gives it back as clean, editable text in Word or another format.

Just ask Claude: *"get the text out of this scanned contract and give me a Word file"*.

## What it handles

| Input | How it is read |
|---|---|
| Digital PDFs | the PDF's own text layer (exact) |
| Scanned / image-only PDFs, photos, screenshots, HEIC, multi-page TIFF | OCR |
| Word, Excel, PowerPoint, ODT, EPUB, HTML, RTF, TXT | read directly, pictures inside can be OCR'd too |

Output: **Word (.docx)**, ODT, RTF, HTML, Markdown, plain text. Headings, lists and tables are kept.

## How it works

1. **Extract.** It uses the text layer when there is one. Otherwise it OCRs with Apple Vision (macOS), or with Tesseract for languages Vision lacks.
2. **Review.** Weak pages are flagged: low confidence, low-resolution photos, tables, or very little text. Claude looks at those page images itself and fixes mistakes, rebuilds tables and reads handwriting. Anything it can't read is marked `[unreadable]`, never guessed.
3. **Export.** The text is saved in the format you asked for.

## Install

### Claude Code (plugin)

In Claude Code, run:

```
/plugin marketplace add fplusf/doc-ocr
/plugin install doc-ocr@doc-ocr
```

### Claude Code (manual)

```bash
git clone https://github.com/fplusf/doc-ocr.git
cp -R doc-ocr/skills/doc-ocr ~/.claude/skills/
```

### Claude.ai

Download `doc-ocr.skill` from [Releases](https://github.com/fplusf/doc-ocr/releases) and upload it under **Settings > Capabilities > Skills**.

## Requirements

- Python 3.9+ with `pymupdf` and `python-docx`:
  ```bash
  pip3 install pymupdf python-docx
  ```
- OCR engine, at least one of:
  - **macOS:** Apple Vision. Built in, and compiled automatically on first use (needs Xcode command line tools).
  - **Any OS:** Tesseract (`brew install tesseract` or `apt install tesseract-ocr`).
  - **Neither:** Claude reads the page images itself (slower, but works).
- Optional: `openpyxl` (Excel), `beautifulsoup4` (HTML/EPUB).

Check your setup:

```bash
bash ~/.claude/skills/doc-ocr/scripts/setup.sh
```

### More languages

Apple Vision covers English, Russian, Ukrainian, most of Western Europe, Turkish, Arabic, Chinese, Japanese, Korean and more. For **Tajik, Uzbek, Kazakh, Kyrgyz, Persian, Hindi** add Tesseract packs:

```bash
bash ~/.claude/skills/doc-ocr/scripts/setup.sh --langs tgk rus
```

## Use the scripts without Claude

```bash
python3 skills/doc-ocr/scripts/extract.py scan.pdf photo.jpg --out-dir out --lang en,ru
python3 skills/doc-ocr/scripts/export.py out/text.md --to docx --out-base result
```

Run `--help` on either script for all options.

## Privacy

Everything runs on your machine. The scripts make no network calls. The only download is the optional language-pack command above.

## License

MIT
