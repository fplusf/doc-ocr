#!/usr/bin/env python3
"""Extract clean text from (almost) any document: scanned or digital PDFs,
photos and screenshots, Word, Excel, PowerPoint, OpenDocument, EPUB, HTML,
plain text. Scanned pages and images go through OCR.

Output (in --out-dir):
  text.md      clean Markdown, one section per page, with page markers
  report.json  per-page engine, confidence and a "needs_review" flag
  pages/       PNG of every page that needs a visual check (or all with --save-pages)

Usage:
  extract.py INPUT [INPUT ...] --out-dir DIR [--lang en,ru] [--engine auto]
             [--force-ocr] [--save-pages] [--keep-lines] [--dpi 300]

Engines:
  auto       native text layer when present, else Apple Vision (macOS), else Tesseract
  native     text layer only, never OCR
  vision     Apple Vision OCR (macOS, best quality for photos and print)
  tesseract  Tesseract OCR (cross-platform; needed for languages Vision lacks, e.g. Tajik)
  claude     no OCR: render every page to PNG so Claude reads the images itself
"""
from __future__ import annotations

import argparse
import html
import io
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
CACHE = Path(os.environ.get("DOC_OCR_CACHE", Path.home() / ".cache" / "doc-ocr"))
VISION_BIN = CACHE / "vision_ocr"

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp", ".heic", ".heif", ".jp2"}
TEXT_EXT = {".txt", ".md", ".markdown", ".csv", ".tsv", ".json", ".xml", ".log"}

# Short language code -> (Apple Vision code, Tesseract code)
LANGS = {
    "en": ("en-US", "eng"), "ru": ("ru-RU", "rus"), "tg": (None, "tgk"), "uz": (None, "uzb"),
    "uk": ("uk-UA", "ukr"), "kk": (None, "kaz"), "ky": (None, "kir"), "fa": (None, "fas"),
    "ar": ("ar-SA", "ara"), "tr": ("tr-TR", "tur"), "de": ("de-DE", "deu"), "fr": ("fr-FR", "fra"),
    "es": ("es-ES", "spa"), "it": ("it-IT", "ita"), "pt": ("pt-BR", "por"), "nl": ("nl-NL", "nld"),
    "pl": ("pl-PL", "pol"), "cs": ("cs-CZ", "ces"), "ro": ("ro-RO", "ron"), "sv": ("sv-SE", "swe"),
    "da": ("da-DK", "dan"), "no": ("nb-NO", "nor"), "id": ("id-ID", "ind"), "ms": ("ms-MY", "msa"),
    "vi": ("vi-VT", "vie"), "th": ("th-TH", "tha"), "ko": ("ko-KR", "kor"), "ja": ("ja-JP", "jpn"),
    "zh": ("zh-Hans", "chi_sim"), "zh-hant": ("zh-Hant", "chi_tra"), "hi": (None, "hin"),
}

# Review thresholds: a page under these is flagged for a visual check by Claude.
VISION_MIN_CONF = 0.80
TESS_MIN_CONF = 0.75


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


# --------------------------------------------------------------------------- engines

def ensure_vision() -> bool:
    if sys.platform != "darwin":
        return False
    src = HERE / "vision_ocr.swift"
    if VISION_BIN.exists() and (VISION_BIN.stat().st_mtime >= src.stat().st_mtime or not shutil.which("swiftc")):
        return True
    if not shutil.which("swiftc"):
        return False
    CACHE.mkdir(parents=True, exist_ok=True)
    log("compiling Apple Vision helper (one time, ~30 s)...")
    r = subprocess.run(["swiftc", "-O", str(src), "-o", str(VISION_BIN)], capture_output=True, text=True)
    if r.returncode != 0:
        log("could not compile Vision helper: " + r.stderr[-500:])
        return False
    return True


def tesseract_langs() -> set[str]:
    if not shutil.which("tesseract"):
        return set()
    r = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True)
    return {l.strip() for l in r.stdout.splitlines()[1:] if l.strip()}


class Engines:
    def __init__(self, langs: list[str], engine: str):
        self.langs = langs
        self.requested = engine
        self.vision_ok = engine in ("auto", "vision") and ensure_vision()
        self.tess_have = tesseract_langs()
        self.vision_langs = [LANGS[l][0] for l in langs if l in LANGS and LANGS[l][0]]
        self.vision_missing = [l for l in langs if l not in LANGS or not LANGS[l][0]]
        self.tess_wanted = [LANGS[l][1] if l in LANGS else l for l in langs]
        self.tess_missing = [t for t in self.tess_wanted if t not in self.tess_have]
        self.warnings: list[str] = []

    def pick(self) -> str:
        if self.requested in ("vision", "tesseract", "claude"):
            return self.requested
        # auto: Vision unless it lacks a requested language that Tesseract has
        tess_usable = self.tess_have and not self.tess_missing
        if self.vision_ok and (not self.vision_missing or not tess_usable):
            if self.vision_missing:
                self.warnings.append(
                    f"Apple Vision has no model for {self.vision_missing}; Tesseract pack missing too. "
                    f"Using Vision with the closest script; install Tesseract packs for better results "
                    f"(see references/setup.md).")
            return "vision"
        if tess_usable:
            return "tesseract"
        if self.tess_have:
            self.warnings.append(f"Tesseract language packs missing: {self.tess_missing}; falling back to 'eng'.")
            return "tesseract"
        return "claude"

    # -- each engine returns (markdown_text, mean_confidence 0..1 or None)

    def ocr(self, png: Path, engine: str, keep_lines: bool) -> tuple[str, float | None]:
        if engine == "vision":
            return self._vision(png, keep_lines)
        if engine == "tesseract":
            return self._tesseract(png, keep_lines)
        return "", None

    def _vision(self, png: Path, keep_lines: bool) -> tuple[str, float | None]:
        langs = ",".join(self.vision_langs) if self.vision_langs else "auto"
        r = subprocess.run([str(VISION_BIN), str(png), langs], capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError("vision: " + r.stderr.strip())
        data = json.loads(r.stdout)
        lines = data["lines"]
        if not lines:
            return "", None
        conf = statistics.fmean(l["conf"] for l in lines)
        return layout_lines(lines, keep_lines), conf

    def _tesseract(self, png: Path, keep_lines: bool) -> tuple[str, float | None]:
        langs = [t for t in self.tess_wanted if t in self.tess_have] or ["eng"]
        r = subprocess.run(["tesseract", str(png), "stdout", "-l", "+".join(langs), "--psm", "3", "tsv"],
                           capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError("tesseract: " + r.stderr.strip()[-300:])
        rows = [l.split("\t") for l in r.stdout.splitlines()[1:]]
        words = [w for w in rows if len(w) == 12 and w[11].strip() and float(w[10]) >= 0]
        if not words:
            return "", None
        conf = statistics.fmean(float(w[10]) for w in words) / 100.0
        page_w = next((float(r[8]) for r in rows if len(r) == 12 and r[0] == "1"), 1.0) or 1.0
        paras: dict[tuple, dict[int, list]] = {}
        for w in words:
            pkey = (int(w[2]), int(w[3]))          # block, paragraph
            lkey = int(w[4])                       # line
            left, width = float(w[6]), float(w[8])
            paras.setdefault(pkey, {}).setdefault(lkey, []).append((w[11], left, left + width))
        out = []
        for _, lines in paras.items():
            txt = [" ".join(t for t, _, _ in ws) for ws in lines.values()]
            widths = [(max(r for _, _, r in ws) - min(l for _, l, _ in ws)) / page_w for ws in lines.values()]
            out.append("\n".join(txt) if keep_lines else join_paragraph(txt, widths))
        return "\n\n".join(out), conf


# --------------------------------------------------------------------------- layout

def join_paragraph(lines: list[str], widths: list[float] | None = None) -> str:
    """Join wrapped lines into prose, keep real line breaks (lists, short lines)."""
    if not lines:
        return ""
    maxw = max(widths) if widths else None
    out = lines[0]
    for i in range(1, len(lines)):
        prev, cur = out, lines[i]
        short_prev = widths is not None and maxw and widths[i - 1] < 0.7 * maxw
        listy = re.match(r"^\s*([-•*▪●]|\d{1,3}[.)]|[a-zа-я][.)])\s", cur, re.I)
        labelled = re.match(r"^[^\W\d][\w .«»\"'/()-]{0,40}:\s", cur)      # "Total: 120", "Покупатель: ..."
        if prev.endswith("-") and not prev.endswith(" -") and cur[:1].islower():
            out = prev[:-1] + cur                  # de-hyphenate
        elif short_prev or listy or labelled or prev.endswith(":"):
            out = prev + "\n" + cur
        else:
            out = prev + " " + cur
    return out


def layout_lines(lines: list[dict], keep_lines: bool) -> str:
    """Turn Vision line boxes (normalised, top-left origin) into paragraphs,
    handling simple two-column pages."""
    # merge fragments that sit on the same visual line
    lines = sorted(lines, key=lambda l: (l["y"] + l["h"] / 2, l["x"]))
    narrow = [l for l in lines if l["w"] < 0.47]
    left = [l for l in narrow if l["x"] + l["w"] <= 0.53]
    right = [l for l in narrow if l["x"] >= 0.47]
    two_col = len(lines) >= 12 and len(narrow) / len(lines) > 0.75 and len(left) > 4 and len(right) > 4
    if two_col:
        cols = [[l for l in lines if l["x"] + l["w"] / 2 < 0.5], [l for l in lines if l["x"] + l["w"] / 2 >= 0.5]]
    else:
        cols = [lines]
    parts = []
    for col in cols:
        rows: list[list[dict]] = []
        for l in sorted(col, key=lambda l: (l["y"], l["x"])):
            cy = l["y"] + l["h"] / 2
            if rows:
                r = rows[-1]
                ry, rh = r[0]["y"], max(x["h"] for x in r)
                if ry <= cy <= ry + rh:
                    r.append(l)
                    continue
            rows.append([l])
        merged = []
        for r in rows:
            r.sort(key=lambda l: l["x"])
            cells = [r[0]["text"]]
            for prev, l in zip(r, r[1:]):
                if l["x"] - (prev["x"] + prev["w"]) > 0.05:     # wide gap: new table cell
                    cells.append(l["text"])
                else:
                    cells[-1] += " " + l["text"]
            merged.append({"text": " ".join(cells), "cells": cells, "y": min(l["y"] for l in r),
                           "h": max(l["h"] for l in r), "x": r[0]["x"], "w": r[-1]["x"] + r[-1]["w"] - r[0]["x"]})
        if not merged:
            continue
        med_h = statistics.median(m["h"] for m in merged)
        # split into table runs (3+ consecutive rows with the same 2+ cell count) and text runs
        segments: list[tuple[str, list[dict]]] = []
        i = 0
        while i < len(merged):
            n = len(merged[i]["cells"])
            j = i
            while j < len(merged) and n >= 2 and len(merged[j]["cells"]) == n:
                j += 1
            if j - i >= 3:
                segments.append(("table", merged[i:j]))
                i = j
            else:
                if segments and segments[-1][0] == "text":
                    segments[-1][1].append(merged[i])
                else:
                    segments.append(("text", [merged[i]]))
                i += 1
        for kind, seg in segments:
            if kind == "table":
                parts.append(md_table([m["cells"] for m in seg]))
                continue
            paras, cur = [], [seg[0]]
            for prev, m in zip(seg, seg[1:]):
                gap = m["y"] - (prev["y"] + prev["h"])
                bigger = m["h"] > 1.35 * med_h or prev["h"] > 1.35 * med_h
                if gap > 0.9 * med_h or bigger or abs(m["x"] - cur[0]["x"]) > 0.15:
                    paras.append(cur)
                    cur = [m]
                else:
                    cur.append(m)
            paras.append(cur)
            for p in paras:
                texts = [m["text"] for m in p]
                body = "\n".join(texts) if keep_lines else join_paragraph(texts, [m["w"] for m in p])
                # a single short line in large type is probably a heading
                if len(p) == 1 and p[0]["h"] > 1.35 * med_h and len(body) < 120:
                    body = "## " + body
                parts.append(body)
    return "\n\n".join(parts)


# --------------------------------------------------------------------------- images

def _upright(src: Path, workdir: Path) -> Path:
    """Convert HEIC and apply the EXIF rotation of phone photos (macOS sips, else Pillow if usable)."""
    ext = src.suffix.lower()
    if ext in (".heic", ".heif") or ext in (".jpg", ".jpeg"):
        if shutil.which("sips"):
            dst = workdir / f"{src.stem}-upright.jpg"
            # sips writes pixels upright and drops the orientation tag when it re-encodes
            r = subprocess.run(["sips", "-s", "format", "jpeg", str(src), "--out", str(dst)], capture_output=True)
            if r.returncode == 0 and dst.exists():
                return dst
        try:
            from PIL import Image, ImageOps  # optional
            dst = workdir / f"{src.stem}-upright.png"
            ImageOps.exif_transpose(Image.open(src)).convert("RGB").save(dst)
            return dst
        except Exception:  # noqa: BLE001
            if ext in (".heic", ".heif"):
                raise RuntimeError("HEIC needs macOS 'sips' or a working Pillow; convert it to JPG first")
    return src


def image_pages(src: Path, workdir: Path) -> list[Path]:
    """Any image (incl. multi-page TIFF) -> list of OCR-ready PNGs, upscaled if small.
    Uses PyMuPDF so it works even when Pillow is missing or broken."""
    import fitz
    src = _upright(src, workdir)
    doc = fitz.open(src)
    out = []
    for i, page in enumerate(doc):
        longest = max(page.rect.width, page.rect.height)
        zoom = max(1.0, 2000 / longest) if longest < 2000 else 1.0   # small screenshots: upscale
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        p = workdir / f"{src.stem}-{i + 1}-prep.png"
        pix.save(p)
        out.append(p)
    return out


def image_size(path: Path) -> tuple[int, int]:
    import fitz
    pix = fitz.Pixmap(str(path))
    return pix.width, pix.height


# --------------------------------------------------------------------------- per-format readers

def looks_garbled(text: str) -> bool:
    """A text layer can exist but be junk (bad font maps, OCR-less scans with fake text)."""
    t = re.sub(r"\s", "", text)
    if not t:
        return True
    bad = sum(1 for c in t if c == "\ufffd" or (ord(c) < 32) or 0xE000 <= ord(c) <= 0xF8FF)
    letters = sum(1 for c in t if c.isalpha())
    return bad / len(t) > 0.05 or letters / len(t) < 0.3


class Job:
    def __init__(self, args):
        self.args = args
        self.out = Path(args.out_dir)
        self.pages_dir = self.out / "pages"
        self.out.mkdir(parents=True, exist_ok=True)
        self.work = Path(tempfile.mkdtemp(prefix="doc-ocr-"))
        self.langs = [l.strip().lower() for l in args.lang.split(",") if l.strip()]
        self.eng = Engines(self.langs, args.engine)
        self.ocr_engine = self.eng.pick() if args.engine != "native" else "native"
        self.report = {"inputs": [], "engine": self.ocr_engine, "languages": self.langs,
                       "warnings": self.eng.warnings, "pages": []}
        self.md: list[str] = []
        self.page_no = 0

    # one "page" record = one PDF page, one image, or one whole non-paged document
    def add_page(self, source: str, label: str, text: str, method: str, conf: float | None,
                 image: Path | None, flag_reason: str | None = None):
        self.page_no += 1
        n = self.page_no
        needs = flag_reason is not None
        if conf is not None:
            limit = VISION_MIN_CONF if method == "vision" else TESS_MIN_CONF
            if conf < limit:
                needs, flag_reason = True, flag_reason or f"low OCR confidence {conf:.2f}"
        if method in ("vision", "tesseract") and len(text.strip()) < 15:
            needs, flag_reason = True, flag_reason or "almost no text found (handwriting, photo, or blank?)"
        if method == "claude":
            needs, flag_reason = True, "no OCR run: Claude must read this image"
        if method in ("vision", "tesseract") and "\n|---" in text:
            needs, flag_reason = True, flag_reason or "table detected: check rows and columns against the image"
        saved = None
        if image is not None and (needs or self.args.save_pages or method in ("vision", "tesseract")):
            self.pages_dir.mkdir(exist_ok=True)
            saved = self.pages_dir / f"page-{n:03d}.png"
            shutil.copy(image, saved)
        conf_s = f"{conf:.2f}" if conf is not None else "n/a"
        self.md.append(f"<!-- page {n} | {source} {label} | {method} | conf {conf_s} -->\n\n{text.strip()}\n")
        self.report["pages"].append({
            "page": n, "source": source, "where": label, "method": method,
            "confidence": None if conf is None else round(conf, 3), "chars": len(text.strip()),
            "needs_review": needs, "reason": flag_reason, "image": str(saved) if saved else None,
        })

    def ocr_png(self, png: Path, source: str, label: str, low_res: bool = False):
        method = self.ocr_engine if self.ocr_engine != "native" else "claude"
        try:
            text, conf = self.eng.ocr(png, method, self.args.keep_lines)
        except Exception as e:  # noqa: BLE001
            log(f"  OCR failed on {source} {label}: {e}")
            text, conf, method = "", None, "claude"
        if any(l in ("ru", "tg", "uk", "kk", "ky", "uz") for l in self.langs):
            text = re.sub(r"\bN[º°o]\s?(?=\d)", "№ ", text)       # Cyrillic number sign
        reason = "low-resolution source image: check every word and number" if low_res else None
        self.add_page(source, label, text, method, conf, png, reason)

    # ---- formats
    def do_pdf(self, path: Path):
        import fitz
        doc = fitz.open(path)
        if doc.needs_pass and not doc.authenticate(self.args.password or ""):
            raise RuntimeError("PDF is password protected: pass --password")
        for i, page in enumerate(doc):
            label = f"p{i + 1}"
            text = page.get_text("text", sort=True)
            has_images = bool(page.get_images(full=False))
            native_ok = len(text.strip()) >= 40 and not looks_garbled(text)
            if self.args.engine == "native" or (native_ok and not self.args.force_ocr):
                blocks = page.get_text("blocks", sort=True)
                paras = [join_paragraph([l for l in b[4].splitlines() if l.strip()]) if not self.args.keep_lines
                         else b[4].strip() for b in blocks if b[6] == 0 and b[4].strip()]
                body = "\n\n".join(paras)
                png = None
                if self.args.save_pages:
                    png = self.work / f"{path.stem}-{i + 1}.png"
                    page.get_pixmap(dpi=110).save(png)
                self.add_page(path.name, label, body, "native", None, png)
                if has_images and self.args.ocr_embedded:
                    self._ocr_pdf_images(doc, page, path.name, label)
                continue
            png = self.work / f"{path.stem}-{i + 1}.png"
            page.get_pixmap(dpi=self.args.dpi).save(png)
            self.ocr_png(png, path.name, label)

    def _ocr_pdf_images(self, doc, page, source, label):
        import fitz
        for j, info in enumerate(page.get_images(full=True)):
            pix = fitz.Pixmap(doc, info[0])
            if pix.width < 300 or pix.height < 120:
                continue                            # logos and icons
            if pix.n - pix.alpha >= 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            p = self.work / f"{Path(source).stem}-{label}-img{j + 1}.png"
            pix.save(p)
            for png in image_pages(p, self.work):
                self.ocr_png(png, source, f"{label} image {j + 1}")

    def do_image(self, path: Path):
        try:
            low = max(image_size(path)) < 1200
        except Exception:  # noqa: BLE001
            low = False
        pngs = image_pages(path, self.work)
        for k, png in enumerate(pngs):
            self.ocr_png(png, path.name, f"frame {k + 1}" if len(pngs) > 1 else "image", low)

    def do_docx(self, path: Path):
        import docx
        d = docx.Document(str(path))
        out = []
        body = d.element.body
        for child in body.iterchildren():
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                p = docx.text.paragraph.Paragraph(child, d)
                t = p.text.strip()
                if not t:
                    continue
                style = (p.style.name or "").lower() if p.style is not None else ""
                m = re.match(r"heading (\d)", style)
                if m:
                    t = "#" * min(int(m.group(1)), 6) + " " + t
                elif "list" in style:
                    t = "- " + t
                out.append(t)
            elif tag == "tbl":
                tbl = docx.table.Table(child, d)
                out.append(md_table([[c.text.strip() for c in r.cells] for r in tbl.rows]))
        self.add_page(path.name, "document", "\n\n".join(out), "native", None, None)
        if self.args.ocr_embedded:
            self._ocr_zip_media(path, "word/media/")

    def _ocr_zip_media(self, path: Path, prefix: str):
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if name.startswith(prefix) and Path(name).suffix.lower() in IMAGE_EXT:
                    p = self.work / Path(name).name
                    p.write_bytes(z.read(name))
                    try:
                        w, h = image_size(p)
                        if w < 300 or h < 120:
                            continue                # logos and icons
                        for png in image_pages(p, self.work):
                            self.ocr_png(png, path.name, f"embedded {Path(name).name}")
                    except Exception as e:  # noqa: BLE001
                        log(f"  skipped embedded image {name}: {e}")

    def do_xlsx(self, path: Path):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        for ws in wb.worksheets:
            rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
            rows = [r for r in rows if any(c.strip() for c in r)]
            self.add_page(path.name, f"sheet {ws.title}", f"## {ws.title}\n\n" + md_table(rows), "native", None, None)

    def do_pptx(self, path: Path):
        with zipfile.ZipFile(path) as z:
            slides = sorted((n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                            key=lambda n: int(re.findall(r"\d+", n)[-1]))
            for s in slides:
                xml = z.read(s).decode("utf8", "ignore")
                paras = []
                for p in re.findall(r"<a:p>(.*?)</a:p>", xml, re.S):
                    t = "".join(re.findall(r"<a:t>(.*?)</a:t>", p, re.S))
                    if t.strip():
                        paras.append(html.unescape(t))
                n = re.findall(r"\d+", s)[-1]
                self.add_page(path.name, f"slide {n}", "\n\n".join(paras), "native", None, None)
        if self.args.ocr_embedded:
            self._ocr_zip_media(path, "ppt/media/")

    def do_odf(self, path: Path):
        with zipfile.ZipFile(path) as z:
            xml = z.read("content.xml").decode("utf8", "ignore")
        paras = []
        for m in re.finditer(r"<text:(h|p)[^>]*>(.*?)</text:\1>", xml, re.S):
            t = html.unescape(re.sub(r"<[^>]+>", "", m.group(2))).strip()
            if t:
                paras.append(("## " if m.group(1) == "h" else "") + t)
        self.add_page(path.name, "document", "\n\n".join(paras), "native", None, None)

    def do_html(self, path: Path, raw: str | None = None):
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(raw if raw is not None else path.read_text("utf8", "ignore"), "html.parser")
        for t in soup(["script", "style", "nav", "noscript"]):
            t.decompose()
        out = []
        for el in soup.find_all(["h1", "h2", "h3", "h4", "p", "li", "pre", "blockquote", "td"]):
            t = el.get_text(" ", strip=True)
            if not t or el.find_parent(["li", "p"]) is not None and el.name != "li":
                continue
            if el.name.startswith("h"):
                t = "#" * int(el.name[1]) + " " + t
            elif el.name == "li":
                t = "- " + t
            out.append(t)
        if not out:
            out = [soup.get_text("\n", strip=True)]
        return "\n\n".join(out)

    def do_epub(self, path: Path):
        with zipfile.ZipFile(path) as z:
            names = sorted(n for n in z.namelist() if n.endswith((".xhtml", ".html", ".htm")))
            for n in names:
                text = self.do_html(path, z.read(n).decode("utf8", "ignore"))
                if text.strip():
                    self.add_page(path.name, Path(n).name, text, "native", None, None)

    def do_rtf(self, path: Path):
        raw = path.read_text("latin-1", "ignore")
        raw = re.sub(r"\\'([0-9a-f]{2})", lambda m: bytes([int(m.group(1), 16)]).decode("cp1251", "ignore"), raw)
        raw = re.sub(r"\\u(-?\d+)\??", lambda m: chr(int(m.group(1)) % 65536), raw)
        raw = re.sub(r"\\par[d]?\b", "\n", raw)
        raw = re.sub(r"\{\\\*[^{}]*\}|\\[a-z]+-?\d* ?|[{}]", "", raw)
        self.add_page(path.name, "document", re.sub(r"\n{3,}", "\n\n", raw).strip(), "native", None, None)

    def run(self, inputs: list[Path]):
        files: list[Path] = []
        for p in inputs:
            files += sorted(x for x in p.rglob("*") if x.is_file()) if p.is_dir() else [p]
        for f in files:
            ext = f.suffix.lower()
            log(f"reading {f.name}")
            self.report["inputs"].append(str(f))
            try:
                if ext == ".pdf":
                    self.do_pdf(f)
                elif ext in IMAGE_EXT:
                    self.do_image(f)
                elif ext in (".docx", ".docm", ".dotx"):
                    self.do_docx(f)
                elif ext in (".xlsx", ".xlsm"):
                    self.do_xlsx(f)
                elif ext == ".pptx":
                    self.do_pptx(f)
                elif ext in (".odt", ".ods", ".odp"):
                    self.do_odf(f)
                elif ext in (".html", ".htm", ".xhtml"):
                    self.add_page(f.name, "document", self.do_html(f), "native", None, None)
                elif ext == ".epub":
                    self.do_epub(f)
                elif ext == ".rtf":
                    self.do_rtf(f)
                elif ext in TEXT_EXT:
                    self.add_page(f.name, "document", f.read_text("utf8", "ignore"), "native", None, None)
                elif ext in (".doc", ".xls", ".ppt"):
                    raise RuntimeError("old binary Office format: save it as .docx/.xlsx/.pptx or PDF first "
                                       "(on macOS: textutil -convert docx file.doc)")
                else:
                    raise RuntimeError(f"unsupported file type {ext}")
            except Exception as e:  # noqa: BLE001
                log(f"  FAILED {f.name}: {e}")
                self.report.setdefault("errors", []).append({"file": str(f), "error": str(e)})
        (self.out / "text.md").write_text("\n".join(self.md), "utf8")
        pages = self.report["pages"]
        self.report["summary"] = {
            "pages": len(pages),
            "ocr_pages": sum(p["method"] in ("vision", "tesseract") for p in pages),
            "needs_review": [p["page"] for p in pages if p["needs_review"]],
            "chars": sum(p["chars"] for p in pages),
        }
        (self.out / "report.json").write_text(json.dumps(self.report, ensure_ascii=False, indent=2), "utf8")
        shutil.rmtree(self.work, ignore_errors=True)
        s = self.report["summary"]
        print(json.dumps({"text": str(self.out / "text.md"), "report": str(self.out / "report.json"),
                          "engine": self.ocr_engine, **s, "warnings": self.report["warnings"],
                          "errors": self.report.get("errors", [])}, ensure_ascii=False, indent=2))


def md_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    w = max(len(r) for r in rows)
    rows = [[c.replace("|", "\\|").replace("\n", " ") for c in r] + [""] * (w - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * w]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", type=Path)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--lang", default="en", help="comma list of short codes, e.g. en,ru or tg,ru")
    ap.add_argument("--engine", default="auto", choices=["auto", "native", "vision", "tesseract", "claude"])
    ap.add_argument("--force-ocr", action="store_true", help="OCR PDF pages even if they have a text layer")
    ap.add_argument("--ocr-embedded", action="store_true", help="also OCR pictures inside PDFs/DOCX/PPTX")
    ap.add_argument("--save-pages", action="store_true", help="save a PNG of every page, not only flagged ones")
    ap.add_argument("--keep-lines", action="store_true", help="keep original line breaks (forms, poems, code)")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--password", default=None)
    a = ap.parse_args()
    for p in a.inputs:
        if not p.exists():
            sys.exit(f"not found: {p}")
    Job(a).run(a.inputs)


if __name__ == "__main__":
    main()
