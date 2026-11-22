---
name: doc-ocr
description: Read text out of any document, including ones that cannot be edited or copied, and export it as clean text to Word (.docx), OpenDocument, RTF, HTML, Markdown or plain text. Handles scanned PDFs, image-only PDFs, photos of pages, phone pictures, screenshots, receipts, forms, multi-page TIFFs, HEIC, plus Word, Excel, PowerPoint, ODT, EPUB and HTML. Uses the PDF text layer when it exists, Apple Vision or Tesseract OCR when it does not, and Claude's own vision to check and fix weak pages, tables and handwriting. Use this whenever someone wants to OCR something, "get the text out of" a PDF or picture, convert a scan or photo to Word, make a document editable, copy text from a locked or scanned file, digitise paperwork, or turn a pile of documents into text, even if they never say "OCR".
---

# doc-ocr: any document to clean, editable text

The job has three stages. Scripts do the mechanical parts; your judgement does the review, which is where most of the quality comes from.

1. **Extract** with `scripts/extract.py` (text layer, else OCR, page by page).
2. **Review** the pages the script flags, by looking at the page images yourself and fixing the Markdown.
3. **Export** with `scripts/export.py` to the format the person wants (Word by default).

`SKILL_DIR` below means the folder this file is in.

## 1. Extract

```bash
python3 SKILL_DIR/scripts/extract.py INPUT [INPUT...] --out-dir OUT --lang en,ru
```

- `INPUT` can be files or folders (folders are read recursively, in name order). Several inputs go into one `text.md`, so they export as one combined document; run once per input if the person wants separate files.
- `--lang`: short codes, most likely first. Ask only if you cannot tell from the file name, the person's language or a quick look. Common: `en`, `ru`, `tg` (Tajik), `uz`, `de`, `fr`, `ar`, `zh`. Mixed documents: list all, e.g. `tg,ru,en`.
- Useful flags:
  - `--force-ocr` when a PDF has a text layer that is wrong (copy-paste gives junk, or the person says the text is broken).
  - `--ocr-embedded` to also read pictures inside PDFs, Word and PowerPoint files (scanned pages pasted into a .docx, screenshots in slides).
  - `--keep-lines` for forms, poems, addresses, code, receipts, invoices: anything where line breaks carry meaning. Without it, wrapped lines are joined into paragraphs, but `Label: value` lines, list items and short lines are kept apart.
  - `--engine tesseract|vision|claude|native` to override the automatic choice (see below).
  - `--password` for protected PDFs, only with a password the person gave you.

It prints a JSON summary and writes:
- `OUT/text.md`: the text, one section per page, each starting with `<!-- page N | source | engine | conf X -->`. Keep these markers while you edit; the exporter removes them or turns them into page breaks.
- `OUT/report.json`: per page method, confidence, `needs_review` and the reason.
- `OUT/pages/page-NNN.png`: an image of every OCR'd page and every flagged page.

**How the engine is picked** (`--engine auto`): PDF pages with a real text layer are read directly (exact, instant). Everything else is OCR'd with Apple Vision on macOS (very good on print, photos and many languages), or Tesseract when Vision lacks a requested language (Tajik, Uzbek, Kazakh, Persian, Hindi) and the Tesseract pack is installed. If neither engine can run, the pages are rendered to images for you to read directly (`claude` engine). Warnings in the summary tell you when a language fell back; pass them on to the person.

## 2. Review (do not skip)

OCR is right most of the time and confidently wrong the rest. The confidence number is about characters, not layout: a page can score 1.00 and still have merged lines or a wrong reading order. Pages are flagged for low confidence, low-resolution source images, detected tables or almost no text. For each page listed in `needs_review`:

1. Open `OUT/pages/page-NNN.png` with the Read tool and compare it with that page's section in `text.md`.
2. Fix what is wrong, directly in `text.md`: misread words and numbers, merged or split paragraphs, lost headings, wrong reading order in columns, missing lines.
3. Tables: rebuild them as Markdown tables matching the image (header row, same columns). OCR often gets cell text right but the structure wrong.
4. Handwriting, stamps, very faint or photographed pages: read the image yourself and write what you see. Mark anything you truly cannot read as `[unreadable]` rather than guessing. Guessed names, amounts and dates are worse than a visible gap, because the person will trust the output.
5. With the `claude` engine every page is flagged and empty: you transcribe each image.

Also skim pages that were not flagged, with extra care for numbers, amounts, dates, IDs and names, since one wrong digit matters more than a typo. For long documents (more than ~20 flagged pages) spot-check a sample, then tell the person how many pages you reviewed.

Keep fixes faithful: correct OCR mistakes, do not rewrite, summarise, translate or "improve" the author's text unless asked. Leave the original language and spelling alone.

Light structure is welcome when it is clearly in the source: `#` headings for titles, `- ` and `1. ` lists, `**bold**` for clearly bold labels. The exporter understands headings, paragraphs, lists, tables, bold, italic, quotes and code blocks.

## 3. Export

```bash
python3 SKILL_DIR/scripts/export.py OUT/text.md --to docx --out-base PATH/NAME [--title "..."] [--page-breaks]
```

- Formats: `docx` (Word, default), `odt`, `rtf`, `html`, `md`, `txt`. Repeat `--to` for several.
- `--page-breaks` keeps the original page boundaries (good for contracts and forms people compare against the paper). Without it the text flows as one document.
- `--font` / `--size` change the base font (default Calibri 11, which covers Latin and Cyrillic).
- Name the output after the source (`contract.pdf` becomes `contract.docx`) and save it next to the source or where the person asked, never only in a temp folder.

## Reporting back

Tell the person, briefly: what was read, how (text layer or OCR), how many pages you checked and fixed, anything marked `[unreadable]`, any language warning, and the path of each exported file.

## When things go wrong

- `ImportError ... incompatible architecture` on a Mac: the Python is universal and a package was built for the other CPU type. Run the script with `arch -arm64 python3 ...` (or `arch -x86_64`); `setup.sh` prints the launch command that works.
- PDF text layer is junk but not detected: rerun with `--force-ocr`.
- Columns read in the wrong order: fix in review; for many pages, try `--engine tesseract` which segments columns differently.
- Old `.doc`/`.xls`/`.ppt`: on macOS run `textutil -convert docx file.doc` first, or ask for PDF/DOCX.
- Missing tools or a language pack: run `bash SKILL_DIR/scripts/setup.sh` and see `references/setup.md`. Installing software or downloading language packs changes the person's machine, so ask before doing it.
