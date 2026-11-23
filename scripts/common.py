#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""common.py — frozen shared helpers for the academic-PPT template adaptation Skill.

Pure standard library. No third-party imports, no network, no ``__pycache__`` writes
(scripts are run with ``python -B``). All JSON is written UTF-8 with ``ensure_ascii=False``.

Owner: lead (see work/CONTRACT.md §3). Freezing this API is what lets the discovery,
adaptation, audit, render and finalize stages be written independently.

Windows and POSIX safe. Nothing here inspects ``/input`` unless asked to.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import os
import re
import struct
import sys
import unicodedata
import zipfile
from xml.etree import ElementTree as ET

# --------------------------------------------------------------------------------------
# namespaces
# --------------------------------------------------------------------------------------

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

RT = {
    "slide": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    "slideLayout": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    "slideMaster": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    "notesSlide": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
    "notesMaster": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
    "theme": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
}

SKILL_NAME = "academic-ppt-template-adaptation"
SKILL_VERSION = "1.0.0"

DECK_DEFAULT_NAME = "slides.pptx"

# readable-size floors (1/100 pt) — below these a slide is *split* instead of squeezed
FONT_FLOOR = {"title": 2400, "body": 1400, "caption": 1200, "table": 1200, "notes": 1000}


# --------------------------------------------------------------------------------------
# time / hashing / io
# --------------------------------------------------------------------------------------

def now_iso() -> str:
    """UTC ISO-8601, second precision, e.g. ``2026-11-23T03:00:00Z``."""
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_dir(path) -> str:
    if path and not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)
    return path


def read_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def _json_text(obj, sort_keys=False, indent=2) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=indent, sort_keys=sort_keys) + "\n"


def write_json(path, obj, sort_keys=False) -> str:
    """Write UTF-8 JSON and return the sha256 of the bytes written."""
    ensure_dir(os.path.dirname(os.path.abspath(path)))
    payload = _json_text(obj, sort_keys=sort_keys).encode("utf-8")
    with open(path, "wb") as fh:
        fh.write(payload)
    return sha256_bytes(payload)


def write_text(path, text) -> str:
    ensure_dir(os.path.dirname(os.path.abspath(path)))
    if isinstance(text, str):
        text = text.encode("utf-8")
    with open(path, "wb") as fh:
        fh.write(text)
    return sha256_bytes(text)


def append_jsonl(path, obj) -> None:
    ensure_dir(os.path.dirname(os.path.abspath(path)))
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(obj, ensure_ascii=False) + "\n")


def read_jsonl(path) -> list:
    out = []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except Exception:
                    out.append({"_unparsed": line})
    except Exception:
        return []
    return out


def relpath_or_abspath(path, base=None) -> str:
    try:
        if base:
            return os.path.relpath(path, base).replace("\\", "/")
        return os.path.abspath(path).replace("\\", "/")
    except Exception:
        return str(path).replace("\\", "/")


def safe_copy(src, dst) -> str:
    """Copy preserving bytes exactly. Skips when dst already has the same sha256."""
    ensure_dir(os.path.dirname(os.path.abspath(dst)))
    src_sha = sha256_file(src)
    if os.path.isfile(dst) and sha256_file(dst) == src_sha:
        return src_sha
    with open(src, "rb") as fh_in, open(dst, "wb") as fh_out:
        while True:
            chunk = fh_in.read(1 << 20)
            if not chunk:
                break
            fh_out.write(chunk)
    if sha256_file(dst) != src_sha:
        raise IOError("copy verification failed for %s -> %s" % (src, dst))
    return src_sha


# --------------------------------------------------------------------------------------
# input discovery
# --------------------------------------------------------------------------------------

_KIND_BY_EXT = {
    ".pptx": "pptx",
    ".ppt": "ppt",
    ".potx": "potx",
    ".pptm": "pptm",
    ".docx": "docx",
    ".doc": "doc",
    ".pdf": "pdf",
    ".md": "md",
    ".markdown": "md",
    ".txt": "txt",
    ".tex": "tex",
    ".xlsx": "xlsx",
    ".xls": "xls",
    ".csv": "csv",
    ".tsv": "csv",
    ".png": "png",
    ".jpg": "jpg",
    ".jpeg": "jpeg",
    ".gif": "gif",
    ".bmp": "bmp",
    ".svg": "svg",
    ".emf": "emf",
    ".wmf": "wmf",
    ".tif": "tiff",
    ".tiff": "tiff",
    ".webp": "webp",
    ".bib": "bib",
    ".json": "json",
    ".yaml": "yaml",
    ".yml": "yaml",
    ".r": "code",
    ".py": "code",
    ".ipynb": "json",
}

_SKIP_DIRS = {"__pycache__", ".git", "node_modules", ".venv", "venv", ".idea", ".vscode"}


def guess_kind(path) -> str:
    return _KIND_BY_EXT.get(os.path.splitext(str(path))[1].lower(), "other")


def scan_inputs(input_dir) -> list:
    """Inventory every readable file under ``input_dir``.

    Returns ``[{'rel','abs','kind','size','sha256'}]`` sorted by rel path.
    Unreadable files are still listed with ``size``/``sha256`` set to ``None``.
    """
    out = []
    if not input_dir or not os.path.isdir(input_dir):
        return out
    for root, dirs, files in os.walk(input_dir):
        dirs[:] = sorted(d for d in dirs if d not in _SKIP_DIRS)
        for name in sorted(files):
            abs_path = os.path.join(root, name)
            rel = relpath_or_abspath(abs_path, input_dir)
            entry = {
                "rel": rel,
                "abs": abs_path,
                "kind": guess_kind(name),
                "size": None,
                "sha256": None,
            }
            try:
                entry["size"] = os.path.getsize(abs_path)
                entry["sha256"] = sha256_file(abs_path)
            except Exception as exc:  # noqa: BLE001 - honest degradation
                entry["error"] = "%s: %s" % (type(exc).__name__, exc)
            out.append(entry)
    return out


# --------------------------------------------------------------------------------------
# XML helpers
# --------------------------------------------------------------------------------------

def localname(tag) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def xml_bytes(elem, declaration=True) -> bytes:
    body = ET.tostring(elem, encoding="unicode")
    if declaration:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body).encode("utf-8")
    return body.encode("utf-8")


def load_zip_xml(zf, name):
    """Parse one part of an open ``zipfile.ZipFile`` (or a path). Returns Element or None."""
    try:
        if isinstance(zf, zipfile.ZipFile):
            data = zf.read(name)
        else:
            with zipfile.ZipFile(zf) as z:
                data = z.read(name)
        return ET.fromstring(data)
    except Exception:
        return None


def qn(prefix_colon_tag: str) -> str:
    """``qn('p:sld')`` -> ``{ns}sld``."""
    prefix, _, tag = prefix_colon_tag.partition(":")
    return "{%s}%s" % (NS[prefix], tag)


# --------------------------------------------------------------------------------------
# units
# --------------------------------------------------------------------------------------

EMU_PER_PT = 12700
EMU_PER_INCH = 914400


def emu_to_pt(emu) -> float:
    try:
        return float(emu) / EMU_PER_PT
    except Exception:
        return 0.0


def pt_to_emu(pt) -> int:
    return int(round(float(pt) * EMU_PER_PT))


def emu_to_in(emu) -> float:
    try:
        return float(emu) / EMU_PER_INCH
    except Exception:
        return 0.0


# --------------------------------------------------------------------------------------
# image geometry (header parsing only — never guesses)
# --------------------------------------------------------------------------------------

_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _png_size(data: bytes):
    if len(data) >= 24 and data[:8] == _PNG_SIG and data[12:16] == b"IHDR":
        w, h = struct.unpack(">II", data[16:24])
        return int(w), int(h)
    return None


def _gif_size(data: bytes):
    if len(data) >= 10 and data[:6] in (b"GIF87a", b"GIF89a"):
        w, h = struct.unpack("<HH", data[6:10])
        return int(w), int(h)
    return None


def _bmp_size(data: bytes):
    if len(data) >= 26 and data[:2] == b"BM":
        w, h = struct.unpack("<ii", data[18:26])
        return int(abs(w)), int(abs(h))
    return None


def _jpeg_size(data: bytes):
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        return None
    i = 2
    n = len(data)
    while i + 9 < n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        if i + 4 > n:
            return None
        seg_len = struct.unpack(">H", data[i + 2:i + 4])[0]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB,
                      0xCD, 0xCE, 0xCF):
            if i + 9 > n:
                return None
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return int(w), int(h)
        i += 2 + seg_len
    return None


def _svg_size(data: bytes):
    try:
        head = data[:4096].decode("utf-8", "replace")
    except Exception:
        return None
    def _num(attr):
        m = re.search(attr + r'\s*=\s*"([0-9.]+)(pt|px|in|cm|mm)?"', head)
        if not m:
            return None
        val = float(m.group(1))
        unit = m.group(2) or "px"
        factor = {"px": 1.0, "pt": 96.0 / 72.0, "in": 96.0, "cm": 96.0 / 2.54, "mm": 96.0 / 25.4}.get(unit, 1.0)
        return val * factor
    w, h = _num("width"), _num("height")
    if w and h:
        return int(round(w)), int(round(h))
    m = re.search(r'viewBox\s*=\s*"\s*[-0-9.]+\s+[-0-9.]+\s+([0-9.]+)\s+([0-9.]+)', head)
    if m:
        return int(round(float(m.group(1)))), int(round(float(m.group(2))))
    return None


def natural_size(path_or_bytes, kind=None):
    """Pixel size from real file headers. Returns ``(w, h)`` or ``None`` — never a guess."""
    if isinstance(path_or_bytes, (bytes, bytearray)):
        data = bytes(path_or_bytes)
    else:
        try:
            with open(path_or_bytes, "rb") as fh:
                data = fh.read(1 << 16)
        except Exception:
            return None
        if kind is None:
            kind = guess_kind(path_or_bytes)
    for probe in (_png_size, _jpeg_size, _gif_size, _bmp_size, _svg_size):
        try:
            got = probe(data)
        except Exception:
            got = None
        if got and got[0] > 0 and got[1] > 0:
            return got
    return None


# --------------------------------------------------------------------------------------
# text metrics (deterministic, no font files needed)
# --------------------------------------------------------------------------------------

# Adobe Helvetica AFM advance widths / 1000 for ASCII 32..126.
_HELV = (
    "278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 "
    "556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 "
    "1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 "
    "667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 "
    "333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 "
    "556 556 333 500 278 556 500 722 500 500 500 334 260 334 584"
)
_WIDTHS = {}
for _i, _w in enumerate(_HELV.split()):
    _WIDTHS[chr(32 + _i)] = float(_w) / 1000.0


def char_width_em(ch: str) -> float:
    """Advance width in em units for one character."""
    if ch in _WIDTHS:
        return _WIDTHS[ch]
    if ch in ("\t",):
        return 2.0
    o = ord(ch)
    if o < 32:
        return 0.0
    east = unicodedata.east_asian_width(ch)
    if east in ("W", "F"):
        return 1.0
    if east == "A":
        # ambiguous: full width in CJK context, half in Latin context — be conservative
        return 1.0
    if o < 128:
        return 0.55
    return 0.6


def text_width_em(text: str) -> float:
    return sum(char_width_em(c) for c in (text or ""))


def _is_cjk(ch: str) -> bool:
    o = ord(ch)
    return (0x2E80 <= o <= 0xA4CF) or (0xAC00 <= o <= 0xD7A3) or (0xF900 <= o <= 0xFAFF) or o >= 0x20000


def estimate_wrapped_lines(text, font_pt, box_w_pt, line_spacing=1.2) -> dict:
    """Greedy wrap estimate. Deterministic; used to *detect overflow*, never to fake a fit.

    ``font_pt`` is in points, ``box_w_pt`` the usable text width in points.
    Returns ``{'lines','max_line_w_pt','height_pt','fits','widest_line'}``.
    """
    font_pt = float(font_pt or 0.0)
    box_w_pt = float(box_w_pt or 0.0)
    line_spacing = float(line_spacing or 1.2)
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")

    if font_pt <= 0 or box_w_pt <= 0:
        return {"lines": 0, "max_line_w_pt": 0.0, "height_pt": 0.0, "fits": None,
                "widest_line": "", "detail": "missing font size or box width"}

    lines = []
    for para in text.split("\n"):
        if para == "":
            lines.append("")
            continue
        # split into tokens: CJK char runs and ASCII words
        tokens = re.findall(r"[\u2E80-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF]|[^\s\u2E80-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF]+|\s+", para)
        cur = ""
        for tok in tokens:
            if tok.strip() == "":
                if cur:
                    cur += " "
                continue
            cand = cur + tok if cur else tok
            if text_width_em(cand) * font_pt <= box_w_pt or not cur:
                cur = cand
            else:
                lines.append(cur)
                cur = tok
        lines.append(cur)

    widths = [text_width_em(ln) * font_pt for ln in lines]
    max_w = max(widths) if widths else 0.0
    height = len(lines) * font_pt * line_spacing
    widest = lines[widths.index(max_w)] if widths else ""
    return {
        "lines": len(lines),
        "max_line_w_pt": round(max_w, 3),
        "height_pt": round(height, 3),
        "fits": bool(max_w <= box_w_pt + 0.5),
        "widest_line": widest[:80],
    }


def capacity_lines(box_h_pt, font_pt, line_spacing=1.2) -> int:
    if not font_pt or font_pt <= 0:
        return 0
    return int(float(box_h_pt) // (float(font_pt) * float(line_spacing or 1.2)))


# --------------------------------------------------------------------------------------
# logging
# --------------------------------------------------------------------------------------

def log(msg, level="info"):
    sys.stderr.write("[skill:%s] %s\n" % (level, msg))
    sys.stderr.flush()


def die(msg, code=2):
    log(msg, "fatal")
    raise SystemExit(code)


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


def norm_key(s: str) -> str:
    """Normalise a name for matching (case/space/punctuation-insensitive)."""
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"[\s_\-.]+", "", s).strip().lower()
