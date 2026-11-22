#!/usr/bin/env python3
"""Export the cleaned Markdown from extract.py to Word and other text formats.

Usage:
  export.py text.md --to docx [--to txt --to html ...] [--out-base NAME]
            [--title "Title"] [--page-breaks] [--font "Calibri"] [--size 11]

Formats: docx (Word), odt (LibreOffice/Google Docs), rtf (opens everywhere),
         html, md, txt. Format is also inferred from --out with an extension.

Supported Markdown: #..###### headings, paragraphs, "- " and "1. " lists,
| tables |, **bold**, *italic*, `code`, ``` code blocks ```, > quotes,
and the <!-- page N ... --> markers written by extract.py (removed, or
turned into page breaks with --page-breaks).
"""
from __future__ import annotations

import argparse
import html
import re
import sys
import zipfile
from pathlib import Path

PAGE_RE = re.compile(r"^<!--\s*page\b.*?-->\s*$")


# --------------------------------------------------------------------------- parse

def parse(md: str, page_breaks: bool) -> list[dict]:
    """Very small Markdown block parser -> list of blocks."""
    blocks: list[dict] = []
    lines = md.splitlines()
    i, first_page = 0, True
    para: list[str] = []

    def flush():
        if para:
            blocks.append({"t": "p", "text": "\n".join(para)})
            para.clear()

    while i < len(lines):
        ln = lines[i]
        s = ln.strip()
        if PAGE_RE.match(s):
            flush()
            if page_breaks and not first_page:
                blocks.append({"t": "pagebreak"})
            first_page = False
            i += 1
            continue
        if s.startswith("<!--"):               # other comments: skip until closed
            flush()
            while i < len(lines) and "-->" not in lines[i]:
                i += 1
            i += 1
            continue
        if not s:
            flush()
            i += 1
            continue
        if s.startswith("```"):
            flush()
            code = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            blocks.append({"t": "code", "text": "\n".join(code)})
            i += 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            flush()
            blocks.append({"t": "h", "level": len(m.group(1)), "text": m.group(2).strip()})
            i += 1
            continue
        if s.startswith("|") and i + 1 < len(lines) and re.match(r"^\|?\s*:?-{2,}", lines[i + 1].strip()):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                r = lines[i].strip()
                if not re.match(r"^\|?[\s:|-]+\|?$", r):
                    cells = re.split(r"(?<!\\)\|", r.strip("|"))
                    rows.append([c.strip().replace("\\|", "|") for c in cells])
                i += 1
            blocks.append({"t": "table", "rows": rows})
            continue
        m = re.match(r"^([-*•]|\d{1,3}[.)])\s+(.*)$", s)
        if m:
            flush()
            ordered = m.group(1)[0].isdigit()
            items = []
            while i < len(lines):
                mm = re.match(r"^\s*([-*•]|\d{1,3}[.)])\s+(.*)$", lines[i])
                if not mm or mm.group(1)[0].isdigit() != ordered:
                    break
                items.append(mm.group(2))
                i += 1
            blocks.append({"t": "ol" if ordered else "ul", "items": items})
            continue
        if s.startswith(">"):
            flush()
            q = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                q.append(lines[i].strip()[1:].strip())
                i += 1
            blocks.append({"t": "quote", "text": "\n".join(q)})
            continue
        para.append(s)
        i += 1
    flush()
    return blocks


INLINE_RE = re.compile(r"(\*\*[^*]+\*\*|__[^_]+__|\*[^*\s][^*]*\*|`[^`]+`)")


def runs(text: str) -> list[tuple[str, set]]:
    """Split inline Markdown into (text, {'b','i','code'}) runs."""
    out = []
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith(("**", "__")) and len(part) > 4:
            out.append((part[2:-2], {"b"}))
        elif part.startswith("`"):
            out.append((part[1:-1], {"code"}))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            out.append((part[1:-1], {"i"}))
        else:
            out.append((part, set()))
    return out


def plain(text: str) -> str:
    return "".join(t for t, _ in runs(text))


# --------------------------------------------------------------------------- writers

def to_docx(blocks, out: Path, title, font, size):
    import docx
    from docx.enum.text import WD_BREAK
    from docx.oxml.ns import qn
    from docx.shared import Pt

    d = docx.Document()
    st = d.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(size)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    if title:
        d.core_properties.title = title
        d.add_heading(title, 0)

    def add_runs(p, text):
        for chunk, fmt in runs(text):
            lines = chunk.split("\n")
            for k, piece in enumerate(lines):
                r = p.add_run(piece)
                r.bold = "b" in fmt or None
                r.italic = "i" in fmt or None
                if "code" in fmt:
                    r.font.name = "Menlo"
                if k < len(lines) - 1:
                    r.add_break()

    for b in blocks:
        t = b["t"]
        if t == "h":
            add_runs(d.add_heading("", min(b["level"], 9)), b["text"])
        elif t == "p":
            add_runs(d.add_paragraph(), b["text"])
        elif t in ("ul", "ol"):
            for it in b["items"]:
                add_runs(d.add_paragraph(style="List Number" if t == "ol" else "List Bullet"), it)
        elif t == "quote":
            add_runs(d.add_paragraph(style="Quote"), b["text"])
        elif t == "code":
            p = d.add_paragraph()
            r = p.add_run(b["text"])
            r.font.name = "Menlo"
            r.font.size = Pt(max(size - 2, 8))
        elif t == "table":
            rows = b["rows"]
            w = max(len(r) for r in rows)
            tbl = d.add_table(rows=len(rows), cols=w)
            tbl.style = "Table Grid"
            for ri, r in enumerate(rows):
                for ci in range(w):
                    cell = tbl.cell(ri, ci)
                    cell.text = ""
                    add_runs(cell.paragraphs[0], r[ci] if ci < len(r) else "")
                    if ri == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
            d.add_paragraph()
        elif t == "pagebreak":
            d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.save(out)


def to_html(blocks, title) -> str:
    def inl(text):
        s = ""
        for chunk, fmt in runs(text):
            c = html.escape(chunk).replace("\n", "<br>")
            if "code" in fmt:
                c = f"<code>{c}</code>"
            if "b" in fmt:
                c = f"<strong>{c}</strong>"
            if "i" in fmt:
                c = f"<em>{c}</em>"
            s += c
        return s

    body = []
    if title:
        body.append(f"<h1>{html.escape(title)}</h1>")
    for b in blocks:
        t = b["t"]
        if t == "h":
            body.append(f"<h{b['level']}>{inl(b['text'])}</h{b['level']}>")
        elif t == "p":
            body.append(f"<p>{inl(b['text'])}</p>")
        elif t in ("ul", "ol"):
            body.append(f"<{t}>" + "".join(f"<li>{inl(x)}</li>" for x in b["items"]) + f"</{t}>")
        elif t == "quote":
            body.append(f"<blockquote>{inl(b['text'])}</blockquote>")
        elif t == "code":
            body.append(f"<pre>{html.escape(b['text'])}</pre>")
        elif t == "table":
            rows = b["rows"]
            h = "".join(f"<th>{inl(c)}</th>" for c in rows[0])
            r = "".join("<tr>" + "".join(f"<td>{inl(c)}</td>" for c in row) + "</tr>" for row in rows[1:])
            body.append(f"<table><thead><tr>{h}</tr></thead><tbody>{r}</tbody></table>")
        elif t == "pagebreak":
            body.append('<div style="page-break-after:always"></div><hr>')
    css = ("body{font:16px/1.6 -apple-system,Segoe UI,Roboto,sans-serif;max-width:46rem;margin:2rem auto;"
           "padding:0 1rem;color:#1d1d1f;background:#fff}table{border-collapse:collapse;margin:1rem 0}"
           "td,th{border:1px solid #ccc;padding:.35rem .6rem;text-align:left}pre{background:#f4f4f4;"
           "padding:.8rem;overflow:auto}blockquote{border-left:3px solid #ccc;margin:0;padding-left:1rem}")
    return (f"<!doctype html><html><head><meta charset='utf-8'><title>{html.escape(title or 'Document')}</title>"
            f"<style>{css}</style></head><body>\n" + "\n".join(body) + "\n</body></html>\n")


def to_txt(blocks, title) -> str:
    out = [title, "=" * len(title), ""] if title else []
    for b in blocks:
        t = b["t"]
        if t == "h":
            out += [plain(b["text"]).upper() if b["level"] <= 2 else plain(b["text"]), ""]
        elif t in ("p", "quote", "code"):
            out += [plain(b["text"]) if t != "code" else b["text"], ""]
        elif t in ("ul", "ol"):
            out += [(f"{k + 1}. " if t == "ol" else "- ") + plain(x) for k, x in enumerate(b["items"])] + [""]
        elif t == "table":
            rows = [[plain(c) for c in r] for r in b["rows"]]
            w = max(len(r) for r in rows)
            rows = [r + [""] * (w - len(r)) for r in rows]
            widths = [max(len(r[c]) for r in rows) for c in range(w)]
            out += ["  ".join(r[c].ljust(widths[c]) for c in range(w)).rstrip() for r in rows] + [""]
        elif t == "pagebreak":
            out += ["\f"]
    return "\n".join(out).rstrip() + "\n"


def to_md(blocks, title) -> str:
    out = [f"# {title}", ""] if title else []
    for b in blocks:
        t = b["t"]
        if t == "h":
            out += ["#" * b["level"] + " " + b["text"], ""]
        elif t == "p":
            out += [b["text"], ""]
        elif t in ("ul", "ol"):
            out += [(f"{k + 1}. " if t == "ol" else "- ") + x for k, x in enumerate(b["items"])] + [""]
        elif t == "quote":
            out += ["> " + l for l in b["text"].splitlines()] + [""]
        elif t == "code":
            out += ["```", b["text"], "```", ""]
        elif t == "table":
            r = b["rows"]
            out += ["| " + " | ".join(r[0]) + " |", "|" + "---|" * len(r[0])]
            out += ["| " + " | ".join(x) + " |" for x in r[1:]] + [""]
        elif t == "pagebreak":
            out += ["---", ""]
    return "\n".join(out).rstrip() + "\n"


def rtf_escape(s: str) -> str:
    o = []
    for ch in s:
        c = ord(ch)
        if ch in "\\{}":
            o.append("\\" + ch)
        elif ch == "\n":
            o.append("\\line ")
        elif c < 128:
            o.append(ch)
        else:
            o.append(f"\\u{c if c < 32768 else c - 65536}?")
    return "".join(o)


def to_rtf(blocks, title, font, size) -> str:
    hs = size * 2
    def inl(text):
        s = ""
        for chunk, fmt in runs(text):
            c = rtf_escape(chunk)
            if "b" in fmt:
                c = "{\\b " + c + "}"
            if "i" in fmt:
                c = "{\\i " + c + "}"
            s += c
        return s
    out = ["{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 " + rtf_escape(font) + ";}}\\fs" + str(hs)]
    if title:
        out.append("{\\pard\\b\\fs" + str(hs + 16) + " " + rtf_escape(title) + "\\par}")
    for b in blocks:
        t = b["t"]
        if t == "h":
            out.append("{\\pard\\sb240\\sa120\\b\\fs" + str(hs + max(0, 12 - 2 * b["level"])) + " "
                       + inl(b["text"]) + "\\par}")
        elif t in ("p", "quote", "code"):
            out.append("{\\pard\\sa160 " + (inl(b["text"]) if t != "code" else rtf_escape(b["text"])) + "\\par}")
        elif t in ("ul", "ol"):
            for k, x in enumerate(b["items"]):
                mark = f"{k + 1}." if t == "ol" else "\\bullet"
                out.append("{\\pard\\li360\\fi-360\\sa60 " + mark + "\\tab " + inl(x) + "\\par}")
        elif t == "table":
            w = max(len(r) for r in b["rows"])
            cw = 9000 // max(w, 1)
            for r in b["rows"]:
                row = "\\trowd\\trgaph80" + "".join(
                    f"\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx{cw * (c + 1)}" for c in range(w))
                row += "".join("\\pard\\intbl " + inl(r[c] if c < len(r) else "") + "\\cell" for c in range(w))
                out.append(row + "\\row")
            out.append("\\pard\\par")
        elif t == "pagebreak":
            out.append("\\page")
    return "\n".join(out) + "\n}"


def to_odt(blocks, out: Path, title):
    from xml.sax.saxutils import escape as x
    def inl(text):
        s = ""
        for chunk, fmt in runs(text):
            c = x(chunk).replace("\n", "<text:line-break/>")
            if "b" in fmt:
                c = f'<text:span text:style-name="B">{c}</text:span>'
            elif "i" in fmt:
                c = f'<text:span text:style-name="I">{c}</text:span>'
            s += c
        return s
    body = []
    if title:
        body.append(f'<text:h text:outline-level="1">{x(title)}</text:h>')
    for b in blocks:
        t = b["t"]
        if t == "h":
            body.append(f'<text:h text:outline-level="{b["level"]}">{inl(b["text"])}</text:h>')
        elif t in ("p", "quote", "code"):
            body.append(f"<text:p>{inl(b['text']) if t != 'code' else x(b['text'])}</text:p>")
        elif t == "ul":
            body.append("<text:list>" + "".join(f"<text:list-item><text:p>{inl(i)}</text:p></text:list-item>"
                                                  for i in b["items"]) + "</text:list>")
        elif t == "ol":
            body += [f"<text:p>{k + 1}. {inl(i)}</text:p>" for k, i in enumerate(b["items"])]
        elif t == "table":
            w = max(len(r) for r in b["rows"])
            rows = "".join("<table:table-row>" + "".join(
                f"<table:table-cell><text:p>{inl(r[c] if c < len(r) else '')}</text:p></table:table-cell>" for c in range(w))
                + "</table:table-row>" for r in b["rows"])
            body.append(f'<table:table><table:table-column table:number-columns-repeated="{w}"/>{rows}</table:table>')
        elif t == "pagebreak":
            body.append('<text:p text:style-name="PB"/>')
    ns = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
          'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
          'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
          'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
          'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"')
    styles = ('<office:automatic-styles>'
              '<style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>'
              '<style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>'
              '<style:style style:name="PB" style:family="paragraph"><style:paragraph-properties fo:break-before="page"/></style:style>'
              '</office:automatic-styles>')
    content = (f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {ns} office:version="1.2">'
               f'{styles}<office:body><office:text>{"".join(body)}</office:text></office:body></office:document-content>')
    manifest = ('<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
                'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
                '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>'
                '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
                '</manifest:manifest>')
    with zipfile.ZipFile(out, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.text")
        z.writestr("content.xml", content, zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/manifest.xml", manifest, zipfile.ZIP_DEFLATED)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path, help="Markdown file (text.md from extract.py, after review)")
    ap.add_argument("--to", action="append", default=[], choices=["docx", "odt", "rtf", "html", "md", "txt"])
    ap.add_argument("--out", type=Path, action="append", default=[], help="explicit output path(s); format from extension")
    ap.add_argument("--out-base", type=Path, help="output path without extension (default: next to input)")
    ap.add_argument("--title")
    ap.add_argument("--page-breaks", action="store_true", help="start each source page on a new page")
    ap.add_argument("--font", default="Calibri")
    ap.add_argument("--size", type=int, default=11)
    a = ap.parse_args()

    blocks = parse(a.input.read_text("utf8"), a.page_breaks)
    targets = [(p.suffix.lstrip(".").lower(), p) for p in a.out]
    base = a.out_base or a.input.with_suffix("")
    targets += [(f, base.with_suffix("." + f)) for f in a.to]
    if not targets:
        targets = [("docx", base.with_suffix(".docx"))]
    written = []
    for fmt, path in targets:
        path.parent.mkdir(parents=True, exist_ok=True)
        if fmt == "docx":
            to_docx(blocks, path, a.title, a.font, a.size)
        elif fmt == "odt":
            to_odt(blocks, path, a.title)
        elif fmt == "rtf":
            path.write_text(to_rtf(blocks, a.title, a.font, a.size), "ascii")
        elif fmt in ("html", "htm"):
            path.write_text(to_html(blocks, a.title), "utf8")
        elif fmt in ("md", "markdown"):
            path.write_text(to_md(blocks, a.title), "utf8")
        elif fmt == "txt":
            path.write_text(to_txt(blocks, a.title), "utf8")
        else:
            sys.exit(f"unknown format: {fmt}")
        written.append(str(path))
    print("\n".join(written))


if __name__ == "__main__":
    main()
