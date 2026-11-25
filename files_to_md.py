#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
files_to_md.py —— 批量把文件夹内的文件转换为 Markdown，并给输出文件名加上当天日期前缀。

用法示例：
    python3 files_to_md.py "/Users/me/Documents/资料"
    python3 files_to_md.py "/Users/me/Documents/资料" -r          # 连同子文件夹
    python3 files_to_md.py "/Users/me/Documents/资料" --in-place  # 结果直接放原文件夹
    python3 files_to_md.py "/Users/me/Documents/资料" -o ~/Desktop/结果
    python3 files_to_md.py "/Users/me/Documents/资料" --date-format "%Y%m%d"
    python3 files_to_md.py "/Users/me/Documents/资料" --dry-run      # 只看计划，不写文件
    python3 files_to_md.py "/Users/me/Documents/资料" --open-result   # 完成后在访达打开结果目录

默认行为（重要）：
    * 只处理该文件夹第一层的文件，不递归（用 -r 开启递归）
    * 结果放在  源文件夹/md_<当天日期>/  里
    * 文件名形如 2026-11-25_原文件名.md
    * 【绝不删除、修改任何原文件】；同名文件默认加 -2、-3，不覆盖
    * 已在别处存在的日期前缀不会重复添加

支持：txt / md / log / 代码文件 / csv / tsv / json / yaml / xml /
      html / docx / doc / rtf / xlsx / xlsm / pptx / pdf / odt / epub 等。
      图片、音视频会生成一个引用原文件的 .md；无法解析的二进制文件跳过并列出原因。

依赖：Python 3.9+，无需 pip 安装任何东西。
      PDF 优先使用随脚本附带的 vendor/pypdf（纯 Python，已内置）；读取失败时
      会自动降级到 PyMuPDF / pdfminer / pdftotext / macOS PDFKit。
      以下可选库装了会自动使用：pandoc、PyMuPDF、pdfminer.six、pdftotext、openpyxl。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import html
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET

# 随脚本附带的纯 Python PDF 库：不需要 pip 安装，也不依赖系统 swiftc / PDFKit
_VENDOR = Path(__file__).resolve().parent / "vendor"
if (_VENDOR / "pypdf").is_dir() and str(_VENDOR) not in sys.path:
    sys.path.insert(0, str(_VENDOR))

# ----------------------------------------------------------------------------
# 通用小工具
# ----------------------------------------------------------------------------

TEXT_EXT = {
    ".txt", ".text", ".md", ".markdown", ".mdown", ".mkd", ".log", ".srt", ".vtt",
    ".rst", ".ini", ".cfg", ".conf", ".properties", ".env", ".tex", ".bib",
}
CODE_LANG = {
    ".py": "python", ".js": "javascript", ".jsx": "javascript", ".ts": "typescript",
    ".tsx": "typescript", ".java": "java", ".c": "c", ".h": "c", ".cpp": "cpp",
    ".cc": "cpp", ".cs": "csharp", ".go": "go", ".rs": "rust", ".rb": "ruby",
    ".php": "php", ".swift": "swift", ".kt": "kotlin", ".scala": "scala",
    ".sh": "bash", ".zsh": "bash", ".bat": "bat", ".ps1": "powershell",
    ".sql": "sql", ".css": "css", ".scss": "scss", ".vue": "vue", ".lua": "lua",
    ".r": "r", ".m": "matlab", ".pl": "perl",
}
FENCE_LANG = {
    ".json": "json", ".jsonl": "json", ".yaml": "yaml", ".yml": "yaml",
    ".toml": "toml", ".xml": "xml", ".svg": "xml", ".graphql": "graphql",
}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".ico"}
MEDIA_EXT = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".mp4", ".mov", ".avi", ".mkv", ".webm"}
BIN_SKIP_EXT = {
    ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".dmg", ".pkg", ".iso",
    ".exe", ".dll", ".so", ".o", ".bin", ".dat", ".db", ".sqlite", ".sqlite3",
    ".psd", ".ai", ".sketch", ".fig", ".woff", ".woff2", ".ttf", ".otf",
    ".class", ".jar", ".pyc", ".xls", ".ppt", ".numbers", ".pages", ".key",
}

DATE_PREFIX_RE = re.compile(r"^(?:\d{4}[-_.]?\d{2}[-_.]?\d{2}|\d{8})[-_ ]")


class UnsupportedFormat(Exception):
    """这类文件本来就不该转成 md，属于正常跳过，不算脚本出错。"""


def local(tag: str) -> str:
    """去掉 XML 命名空间，'{ns}p' -> 'p'。"""
    return tag.rsplit("}", 1)[-1]


def attr(el: ET.Element, name: str) -> str | None:
    for k, v in el.attrib.items():
        if local(k) == name:
            return v
    return None


def read_text(path: Path) -> str:
    """多编码兜底读取文本。"""
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "big5", "utf-16", "latin-1"):
        try:
            text = data.decode(enc)
            if "\x00" in text:
                continue
            return text
        except (UnicodeDecodeError, UnicodeError):
            continue
    return ""


def clean_md_text(s: str) -> str:
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n{3,}", "\n\n", s).strip() + "\n"


def md_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    def esc(v: object) -> str:
        return str(v if v is not None else "").replace("|", "\\|").replace("\n", "<br>").strip()
    rows = [list(r) + [""] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(map(esc, rows[0])) + " |",
           "| " + " | ".join(["---"] * width) + " |"]
    out += ["| " + " | ".join(map(esc, r)) + " |" for r in rows[1:]]
    return "\n".join(out)


def has_module(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def tool_exists(name: str) -> str | None:
    return shutil.which(name)


# ----------------------------------------------------------------------------
# 各类文件 -> Markdown
# ----------------------------------------------------------------------------

def from_plain(path: Path, ext: str) -> str:
    text = read_text(path)
    if ext in (".md", ".markdown", ".mdown", ".mkd"):
        return clean_md_text(text)
    if ext in CODE_LANG:
        return f"```{CODE_LANG[ext]}\n{text.strip()}\n```\n"
    if ext in FENCE_LANG:
        if ext == ".json":
            try:
                text = json.dumps(json.loads(text), ensure_ascii=False, indent=2)
            except (ValueError, TypeError):
                pass
        return f"```{FENCE_LANG[ext]}\n{text.strip()}\n```\n"
    return clean_md_text(text)


def from_csv(path: Path, ext: str) -> str:
    text = read_text(path)
    delimiter = "\t" if ext == ".tsv" else ","
    try:
        rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    except csv.Error:
        rows = [line.split(delimiter) for line in text.splitlines()]
    return md_table(rows) + "\n"


class Html2Md(HTMLParser):
    """轻量 HTML -> Markdown，覆盖常见标签，不依赖第三方库。"""

    SKIP = {"script", "style", "noscript", "svg", "iframe", "canvas"}
    VOID = {"br", "hr", "img", "input", "meta", "link", "source", "area", "wbr"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self.skip = 0
        self.pre = 0
        self.href: str | None = None
        self.link_buf: list[str] | None = None
        self.lists: list[list] = []          # [kind, counter]
        self.rows: list[list[str]] = []
        self.row: list[str] | None = None
        self.cell: list[str] | None = None
        self.in_table = False
        self.title_mode = False
        self.title_buf: list[str] = []

    def emit(self, s: str) -> None:
        if self.skip:
            return
        if self.cell is not None:
            self.cell.append(s)
        elif self.link_buf is not None:
            self.link_buf.append(s)
        elif self.title_mode:
            self.title_buf.append(s)
        else:
            self.out.append(s)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        a = dict(attrs)
        if tag in self.SKIP and tag not in self.VOID:
            self.skip += 1
            return
        if self.skip:
            return
        if tag == "title":
            self.title_mode = True
            return
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self.emit("\n" + "#" * int(tag[1]) + " ")
        elif tag == "p":
            self.emit("\n")
        elif tag == "div":
            self.emit("\n")
        elif tag == "br":
            self.emit("  \n")
        elif tag == "hr":
            self.emit("\n\n---\n\n")
        elif tag in ("strong", "b"):
            self.emit("**")
        elif tag in ("em", "i"):
            self.emit("*")
        elif tag == "code" and not self.pre:
            self.emit("`")
        elif tag == "pre":
            self.pre += 1
            self.emit("\n```\n")
        elif tag == "blockquote":
            self.emit("\n> ")
        elif tag in ("ul", "ol"):
            self.lists.append(["ul" if tag == "ul" else "ol", 0])
            self.emit("\n")
        elif tag == "li" and self.lists:
            kind = self.lists[-1][0]
            if kind == "ol":
                self.lists[-1][1] += 1
                self.emit(f"{self.lists[-1][1]}. ")
            else:
                self.emit("- ")
        elif tag == "a":
            self.href = a.get("href") or ""
            self.link_buf = []
        elif tag == "img":
            alt = a.get("alt") or ""
            src = a.get("src") or ""
            self.emit(f"![{alt}]({src})" if src else f"![{alt}]")
        elif tag == "table":
            self.in_table, self.rows = True, []
        elif tag == "tr" and self.in_table:
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []
        elif tag == "head":
            pass

    def handle_endtag(self, tag: str) -> None:
        if tag in self.SKIP and tag not in self.VOID:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag == "title":
            self.title_mode = False
            title = "".join(self.title_buf).strip()
            if title:
                self.out.insert(0, "# " + title + "\n\n")
            return
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6", "p", "div", "li"):
            self.emit("\n")
        elif tag in ("strong", "b"):
            self.emit("**")
        elif tag in ("em", "i"):
            self.emit("*")
        elif tag == "code" and not self.pre:
            self.emit("`")
        elif tag == "pre":
            self.emit("\n```\n")
            self.pre = max(0, self.pre - 1)
        elif tag in ("ul", "ol"):
            if self.lists:
                self.lists.pop()
            self.emit("\n")
        elif tag == "a":
            text = "".join(self.link_buf or []).strip()
            href = self.href or ""
            self.link_buf = None
            self.href = None
            self.emit(f"[{text}]({href})" if href else text)
        elif tag in ("td", "th") and self.cell is not None and self.row is not None:
            self.row.append("".join(self.cell).strip())
            self.cell = None
        elif tag == "tr" and self.row is not None:
            self.rows.append(self.row)
            self.row = None
        elif tag == "table" and self.in_table:
            self.in_table = False
            if self.rows:
                self.out.append("\n\n" + md_table(self.rows) + "\n\n")
            self.rows = []

    def handle_data(self, data: str) -> None:
        if self.pre:
            self.emit(data)
        else:
            self.emit(data)

    def result(self) -> str:
        return clean_md_text(html.unescape("".join(self.out)))


def from_html(path: Path) -> str:
    parser = Html2Md()
    parser.feed(read_text(path))
    parser.close()
    return parser.result()


def para_text(p: ET.Element) -> str:
    buf: list[str] = []
    for node in p.iter():
        tag = local(node.tag)
        if tag == "t" and node.text:
            buf.append(node.text)
        elif tag == "tab":
            buf.append("\t")
        elif tag in ("br", "cr"):
            buf.append("  \n")
    return "".join(buf)


def from_docx(path: Path) -> str:
    """word/document.xml：段落、标题、列表、表格。"""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = next((el for el in root if local(el.tag) == "body"), root)
    lines: list[str] = []
    for el in body:
        tag = local(el.tag)
        if tag == "p":
            text = para_text(el).strip()
            ppr = next((c for c in el if local(c.tag) == "pPr"), None)
            level = 0
            if ppr is not None:
                style = next((c for c in ppr if local(c.tag) == "pStyle"), None)
                sv = (attr(style, "val") or "").lower() if style is not None else ""
                m = re.search(r"(?:heading|标题\s*)(\d)", sv)
                if m:
                    level = min(int(m.group(1)), 6)
                elif sv.isdigit() and 1 <= int(sv) <= 6:
                    level = int(sv)
                elif "title" in sv:
                    level = 1
            if level and text:
                lines.append("#" * level + " " + text)
            elif text:
                numbered = ppr is not None and any(local(c.tag) == "numPr" for c in ppr)
                lines.append(("- " if numbered else "") + text)
            else:
                lines.append("")
        elif tag == "tbl":
            rows: list[list[str]] = []
            for tr in (c for c in el if local(c.tag) == "tr"):
                row = []
                for tc in (c for c in tr if local(c.tag) == "tc"):
                    row.append(" ".join(para_text(p).strip() for p in tc.iter()
                                       if local(p.tag) == "p").strip())
                rows.append(row)
            if rows:
                lines.append(md_table(rows))
    return clean_md_text("\n\n".join(lines))


def col_index(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref.upper())
    n = 0
    for ch in (m.group(1) if m else "A"):
        n = n * 26 + ord(ch) - 64
    return max(n - 1, 0)


def from_xlsx_stdlib(path: Path) -> str:
    """不依赖 openpyxl，直接读 xlsx 的 XML。"""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            sr = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sr:
                shared.append("".join(t.text or "" for t in si.iter() if local(t.tag) == "t"))
        sheet_titles: list[str] = []
        if "xl/workbook.xml" in names:
            wr = ET.fromstring(z.read("xl/workbook.xml"))
            sheet_titles = [attr(s, "name") or f"Sheet{i + 1}"
                            for i, s in enumerate(wr.iter()) if local(s.tag) == "sheet"]
        sheet_files = sorted(
            (n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
            key=lambda n: int(re.search(r"sheet(\d+)", n).group(1)),
        )
        sections: list[str] = []
        for i, sf in enumerate(sheet_files):
            root = ET.fromstring(z.read(sf))
            rows: list[list[str]] = []
            for row_el in root.iter():
                if local(row_el.tag) != "row":       # xlsx 里行标签是 row，不是 r
                    continue
                cells: dict[int, str] = {}
                for c in row_el:
                    if local(c.tag) != "c":
                        continue
                    idx = col_index(attr(c, "r") or "A")
                    ctype = attr(c, "t")
                    if ctype == "inlineStr":
                        value = "".join(t.text or "" for t in c.iter() if local(t.tag) == "t")
                    else:
                        v = next((x.text or "" for x in c if local(x.tag) == "v"), "")
                        if ctype == "s":
                            try:
                                value = shared[int(v)]
                            except (ValueError, IndexError):
                                value = v
                        elif ctype == "b":
                            value = "TRUE" if v == "1" else "FALSE"
                        else:
                            value = v
                    cells[idx] = value
                width = max(cells.keys(), default=-1) + 1
                rows.append([cells.get(j, "") for j in range(width)])
            while rows and not any(r.strip() for r in rows[-1]):
                rows.pop()
            title = sheet_titles[i] if i < len(sheet_titles) else f"Sheet{i + 1}"
            body = md_table(rows) if rows else "*(空工作表)*"
            sections.append(f"## {title}\n\n{body}")
    return clean_md_text("\n\n".join(sections))


def from_xlsx(path: Path) -> str:
    if has_module("openpyxl"):
        import openpyxl  # type: ignore
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sections = []
        for ws in wb.worksheets:
            rows = [[("" if v is None else str(v)) for v in row]
                    for row in ws.iter_rows(values_only=True)]
            while rows and not any(str(v).strip() for v in rows[-1]):
                rows.pop()
            body = md_table(rows) if rows else "*(空工作表)*"
            sections.append(f"## {ws.title}\n\n{body}")
        wb.close()
        return clean_md_text("\n\n".join(sections))
    return from_xlsx_stdlib(path)


def from_pptx(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        slides = sorted(
            (n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"slide(\d+)", n).group(1)),
        )
        sections = []
        for i, name in enumerate(slides, 1):
            root = ET.fromstring(z.read(name))
            paras = []
            for p in root.iter():
                if local(p.tag) == "p":
                    t = "".join(n.text or "" for n in p.iter() if local(n.tag) == "t").strip()
                    if t:
                        paras.append(t)
            sections.append(f"## 幻灯片 {i}\n\n" + ("\n\n".join(paras) if paras else "*(本页无文字)*"))
    return clean_md_text("\n\n".join(sections))


def from_odt_epub(path: Path, ext: str) -> str:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if ext == ".odt":
            root = ET.fromstring(z.read("content.xml"))
            lines = []
            for p in root.iter():
                tag = local(p.tag)
                if tag in ("p", "h"):
                    text = "".join(p.itertext()).strip()
                    if text:
                        lines.append(("# " if tag == "h" else "") + text)
            return clean_md_text("\n\n".join(lines))
        sections = []
        for name in sorted(n for n in names if n.lower().endswith((".html", ".xhtml", ".htm"))):
            parser = Html2Md()
            parser.feed(z.read(name).decode("utf-8", "replace"))
            sections.append(parser.result())
        return clean_md_text("\n\n".join(sections))


SWIFT_PDFKIT = r'''
import Foundation
import PDFKit
let p = CommandLine.arguments[1]
guard let doc = PDFDocument(url: URL(fileURLWithPath: p)) else { exit(2) }
var pages: [String] = []
for i in 0..<doc.pageCount { pages.append(doc.page(at: i)?.string ?? "") }
print(pages.joined(separator: "\u{1F}"))
'''

_pdf_helper_cache: str | None | bool = None  # None=未尝试, False=不可用


def pdfkit_helper() -> str | bool:
    """编译（并缓存）一个调用 macOS PDFKit 提取 PDF 文字的小助手。"""
    global _pdf_helper_cache
    if _pdf_helper_cache is not False and _pdf_helper_cache is not None:
        return _pdf_helper_cache
    swiftc = tool_exists("swiftc")
    if not swiftc:
        _pdf_helper_cache = False
        return False
    cache = Path.home() / ".cache" / "files_to_md"
    src, bin_ = cache / "pdfkit_text.swift", cache / "pdfkit_text"
    try:
        cache.mkdir(parents=True, exist_ok=True)
        if not bin_.exists() or not src.exists() or src.read_text() != SWIFT_PDFKIT:
            src.write_text(SWIFT_PDFKIT, encoding="utf-8")
            r = subprocess.run([swiftc, str(src), "-o", str(bin_)],
                               capture_output=True, text=True, timeout=180)
            if r.returncode != 0 or not bin_.exists():
                _pdf_helper_cache = False
                return False
        _pdf_helper_cache = str(bin_)
        return _pdf_helper_cache
    except (OSError, subprocess.SubprocessError):
        _pdf_helper_cache = False
        return False


def tidy_pdf_text(s: str) -> str:
    """去掉 PDF 提取常见的填充空格、行尾空白和整页空行。"""
    lines = [ln.rstrip() for ln in s.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return "\n".join(ln for ln in lines if ln.strip())


def pdf_pages(path: Path) -> list[str]:
    """多个 PDF 解析器依次自动降级：内置 pypdf -> 系统库 -> macOS PDFKit。"""
    def by_pypdf() -> list[str]:
        from pypdf import PdfReader  # type: ignore
        out = []
        for page in PdfReader(str(path)).pages:
            try:
                out.append(page.extract_text(extraction_mode="layout") or "")
            except (TypeError, ValueError):       # 老版本 pypdf 没有 layout 模式
                out.append(page.extract_text() or "")
        return out

    def by_fitz() -> list[str]:
        import fitz  # type: ignore
        return [page.get_text() for page in fitz.open(str(path))]

    def by_pdfminer() -> list[str]:
        from pdfminer.high_level import extract_text  # type: ignore
        return extract_text(str(path)).split("\f")

    def by_pdftotext() -> list[str]:
        r = subprocess.run(["pdftotext", "-layout", str(path), "-"],
                            capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            raise RuntimeError("pdftotext 执行失败")
        return r.stdout.split("\f")

    def by_pdfkit() -> list[str]:
        helper = pdfkit_helper()
        if helper:
            r = subprocess.run([str(helper), str(path)], capture_output=True, timeout=120)
            if r.returncode == 0:
                return r.stdout.decode("utf-8", "replace").rstrip("\n").split("\x1f")
        if tool_exists("swift"):
            r = subprocess.run(["swift", "-e", SWIFT_PDFKIT, str(path)],
                               capture_output=True, timeout=180)
            if r.returncode == 0:
                return r.stdout.decode("utf-8", "replace").rstrip("\n").split("\x1f")
        raise RuntimeError("swift / PDFKit 不可用")

    chain = []
    if has_module("pypdf"):
        chain.append(("内置 pypdf", by_pypdf))
    if has_module("fitz"):
        chain.append(("PyMuPDF", by_fitz))
    if has_module("pdfminer"):
        chain.append(("pdfminer", by_pdfminer))
    if tool_exists("pdftotext"):
        chain.append(("pdftotext", by_pdftotext))
    chain.append(("macOS PDFKit", by_pdfkit))

    errors, got_empty = [], False
    for name, fn in chain:
        try:
            pages = fn()
            if any(p.strip() for p in pages):
                return [tidy_pdf_text(p) for p in pages]
            got_empty = True
        except Exception as exc:                  # noqa: BLE001 - 自动换下一个解析器
            errors.append(f"{name}: {exc}")
    if got_empty:
        return [""]     # 能打开 PDF 但没有文字层，交给 from_pdf 提示需要 OCR
    raise RuntimeError("所有 PDF 解析器都失败了（" + "；".join(errors) + "）")


def from_pdf(path: Path) -> str:
    pages = list(pdf_pages(path))
    while pages and not pages[-1].strip():        # 去掉结尾的空页
        pages.pop()
    if not pages or not any(p.strip() for p in pages):
        raise UnsupportedFormat("未提取到文字，可能是扫描版/图片 PDF（需要 OCR）")
    if len(pages) == 1:
        return clean_md_text(pages[0])
    sections = [f"## 第 {i} 页\n\n{clean_md_text(p)}" for i, p in enumerate(pages, 1)]
    return clean_md_text("\n\n".join(sections))


def from_textutil(path: Path) -> str:
    exe = tool_exists("textutil")
    if not exe:
        raise RuntimeError("需要 macOS 的 textutil（或把 .doc 另存为 .docx）")
    r = subprocess.run([exe, "-convert", "txt", "-stdout", str(path)],
                       capture_output=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError("textutil 无法读取该文件")
    return clean_md_text(r.stdout.decode("utf-8", "replace"))


def from_media(path: Path, out_dir: Path, ext: str) -> str:
    rel = os.path.relpath(path, out_dir).replace(os.sep, "/")
    label = path.name
    if ext in IMAGE_EXT:
        embed = f"![{label}]({rel})"
    else:
        embed = f"[{label}]({rel})"
    return (f"# {label}\n\n{embed}\n\n"
            f"- 原始文件：`{path}`\n"
            f"- 类型：{'图片' if ext in IMAGE_EXT else '音视频'}\n")


# ----------------------------------------------------------------------------
# 分派 + 主流程
# ----------------------------------------------------------------------------

def detect_kind(path: Path, ext: str) -> str:
    """返回office/pdf/...；对改了扩展名的文件按文件头识别。"""
    try:
        head = path.open("rb").read(4)
    except OSError:
        head = b""
    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as z:
                names = set(z.namelist())
            if "word/document.xml" in names:
                return "docx"
            if "xl/workbook.xml" in names:
                return "xlsx"
            if any(n.startswith("ppt/slides/") for n in names):
                return "pptx"
            if "content.xml" in names:
                return "odt"
            if "mimetype" in names:
                return "epub"
        except (zipfile.BadZipFile, OSError):
            pass
    return ext


def convert(path: Path, out_dir: Path, media_links: bool) -> str:
    ext = path.suffix.lower()
    kind = detect_kind(path, ext)
    if kind in (".docx", "docx"):
        return from_docx(path)
    if kind in (".xlsx", ".xlsm", "xlsx"):
        return from_xlsx(path)
    if kind in (".pptx", "pptx"):
        return from_pptx(path)
    if kind in (".pdf", "pdf"):
        return from_pdf(path)
    if kind in (".odt", ".epub", "odt", "epub"):
        return from_odt_epub(path, ext if ext in (".odt", ".epub") else kind.lstrip("."))
    if kind in (".html", ".htm", ".xhtml", ".shtml"):
        return from_html(path)
    if kind in (".csv", ".tsv"):
        return from_csv(path, ext)
    if kind in (".doc", ".rtf", ".rtfd"):
        return from_textutil(path)
    if ext in IMAGE_EXT or ext in MEDIA_EXT:
        if not media_links:
            raise UnsupportedFormat("已设置不生成图片/音视频引用")
        return from_media(path, out_dir, ext)
    if ext in BIN_SKIP_EXT:
        hint = "（旧版 .doc/.xls/.ppt 建议先用 Office/WPS 另存为新格式）" if ext in (".xls", ".ppt") else ""
        raise UnsupportedFormat(f"暂不支持的二进制格式 {ext}{hint}")
    data = path.read_bytes()
    if not data:
        return ""
    # 未知扩展名：能按文本解码就当文本，否则跳过
    if b"\x00" in data[:4096]:
        raise UnsupportedFormat("二进制文件，无法转为文本")
    return from_plain(path, ext)


def collect_files(folder: Path, recursive: bool, include_hidden: bool,
                  out_dir: Path) -> list[Path]:
    it = folder.rglob("*") if recursive else folder.glob("*")
    out_resolved = out_dir.resolve()
    folder_resolved = folder.resolve()
    files = []
    for p in sorted(it):
        if not p.is_file() or p.is_symlink():
            continue
        if not include_hidden and (p.name.startswith(".") or
                                  any(part.startswith(".") for part in p.relative_to(folder).parts[:-1])):
            continue
        if p.name in (".DS_Store", "Thumbs.db") or p.name.startswith("~$"):
            continue
        try:
            # 输出目录在源文件夹内时，跳过上次生成的结果；--in-place 时两者相同，不能跳过任何文件
            if out_resolved != folder_resolved and out_resolved in p.resolve().parents:
                continue
        except OSError:
            pass
        files.append(p)
    return files


def safe_stem(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"[/\\\x00-\x1f]", "_", stem).strip() or "未命名"


def target_path(out_dir: Path, src: Path, date_str: str,
                overwrite: bool) -> tuple[Path, str]:
    """返回 (实际写入路径, 理想文件名)。"""
    stem = safe_stem(src.name)
    prefix = "" if DATE_PREFIX_RE.match(stem) else f"{date_str}_"
    stem = stem[:max(10, 195 - len(prefix))]      # 截断时也要保住 .md 后缀
    base = f"{prefix}{stem}.md"
    dest = out_dir / base
    # 若目标就是源文件本身（--in-place 且文件名已合规），交给主流程跳过，不能加 -2
    if dest.resolve() == src.resolve():
        return dest, base
    if overwrite or not dest.exists():
        return dest, base
    n = 2
    while True:
        cand = out_dir / f"{Path(base).stem}-{n}.md"
        if not cand.exists():
            return cand, base
        n += 1


def identical_result(out_dir: Path, base: str, content: str) -> Path | None:
    """重跑时若已有内容完全相同的 md，就不再生成 -2 副本。"""
    stem = Path(base).stem
    pattern = re.compile(re.escape(stem) + r"(?:-\d+)?\.md")
    for cand in sorted(out_dir.glob("*.md")):
        if pattern.fullmatch(cand.name):
            try:
                if cand.read_text(encoding="utf-8") == content:
                    return cand
            except (OSError, UnicodeDecodeError):
                continue
    return None


def main() -> int:
    ap = argparse.ArgumentParser(
        description="把文件夹里的文件转成 Markdown，输出文件名前加当天日期。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：python3 files_to_md.py ~/Desktop/资料 -r --in-place")
    ap.add_argument("folder", nargs="?", help="要处理的文件夹路径（不填则运行时提示输入，可把文件夹直接拖进终端）")
    ap.add_argument("-r", "--recursive", action="store_true", help="包含子文件夹里的文件")
    ap.add_argument("-o", "--output", help="输出目录（默认：源文件夹/md_<日期>）")
    ap.add_argument("--in-place", action="store_true", help="结果直接放在源文件夹里")
    ap.add_argument("--date-format", default="%Y-%m-%d",
                    help="日期格式，默认 %%Y-%%m-%%d（如 2026-11-25）")
    ap.add_argument("--overwrite", action="store_true", help="同名时覆盖（默认自动加 -2）")
    ap.add_argument("--include-hidden", action="store_true", help="包含隐藏文件")
    ap.add_argument("--no-media-links", action="store_true", help="不要为图片/音视频生成引用 md")
    ap.add_argument("--dry-run", action="store_true", help="只列出计划，不写任何文件")
    ap.add_argument("--open-result", action="store_true",
                    help="转换完成后在访达中打开结果目录（仅 macOS）")
    args = ap.parse_args()

    if not args.folder:
        try:
            args.folder = input("请输入文件夹路径（可把文件夹拖进终端后回车）：").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n已取消。")
            return 0
        args.folder = args.folder.strip("'\"").strip()
        if not args.folder:
            print("[失败] 未提供文件夹路径", file=sys.stderr)
            return 2
    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        print(f"[失败] 不是文件夹或不存在：{folder}", file=sys.stderr)
        return 2
    if args.output and args.in_place:
        print("[失败] -o 与 --in-place 不能同时使用", file=sys.stderr)
        return 2

    try:
        date_str = dt.date.today().strftime(args.date_format)
    except (ValueError, TypeError):
        print("[失败] --date-format 无效", file=sys.stderr)
        return 2

    out_dir = (folder if args.in_place else
               Path(args.output).expanduser().resolve() if args.output else
               folder / f"md_{date_str}")
    files = collect_files(folder, args.recursive, args.include_hidden, out_dir)
    if not files:
        print(f"[提示] {folder} 里没有可处理的文件")
        return 0

    print(f"源文件夹：{folder}")
    print(f"输出目录：{out_dir}")
    print(f"日期前缀：{date_str}   共发现 {len(files)} 个文件"
          + ("   [预览模式]" if args.dry_run else ""))
    print("-" * 68)
    if not args.dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    ok, skipped, failed = 0, [], []
    for src in files:
        dest, base = target_path(out_dir, src, date_str, args.overwrite)
        if dest.resolve() == src.resolve():
            skipped.append((src.name, "已经是目标文件"))
            print(f"[跳过] {src.name}  （已经是目标文件）")
            continue
        if args.dry_run:
            note = "  （已存在同名结果，内容相同则跳过，不同则自动加序号）" if dest.name != base else ""
            print(f"[计划] {src.name}  ->  {base}{note}")
            continue
        try:
            content = convert(src, out_dir, not args.no_media_links)
            if dest.name != base and not args.overwrite:
                same = identical_result(out_dir, base, content)
                if same is not None:
                    skipped.append((src.name, f"已有相同结果：{same.name}"))
                    print(f"[跳过] {src.name}  ：已有相同结果 {same.name}")
                    continue
            # Path.write_text 的 newline 参数要 Python 3.10+；用 open() 兼容系统自带的 3.9
            with dest.open("w", encoding="utf-8", newline="\n") as fh:
                fh.write(content)
            ok += 1
            print(f"[完成] {src.name}  ->  {dest.name}  ({len(content)} 字符)")
        except UnsupportedFormat as exc:                # 预期内的跳过
            skipped.append((src.name, str(exc)))
            print(f"[跳过] {src.name}  ：{exc}")
        except Exception as exc:                        # noqa: BLE001 - 单文件失败不影响整批
            failed.append((src.name, str(exc)))
            print(f"[失败] {src.name}  ：{exc}")

    print("-" * 68)
    if args.dry_run:
        print(f"预览完成：共 {len(files)} 个文件，未写入任何内容。")
        return 0
    print(f"成功 {ok} 个；跳过 {len(skipped)} 个；失败 {len(failed)} 个")
    if skipped:
        print("跳过明细：")
        for name, why in skipped[:20]:
            print(f"  - {name}: {why}")
        if len(skipped) > 20:
            print(f"  ... 另有 {len(skipped) - 20} 项")
    if failed:
        print("失败明细：")
        for name, why in failed:
            print(f"  - {name}: {why}")
        if any("pdf" in (name + why).lower() for name, why in failed):
            print("提示：PDF 若缺少解析器，可执行 pip install pypdf 后重试。")
    print(f"结果目录：{out_dir}")
    print("原文件均未改动。")
    if args.open_result and tool_exists("open"):
        subprocess.run(["open", str(out_dir)], check=False)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
