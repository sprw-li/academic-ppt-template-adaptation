#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_pptx.py — independent fidelity/format audit for the adapted deck.

The verifier never trusts the engine's self-report: every claim below is recomputed
from the real pptx bytes (zip + XML) and the frozen ``source_inventory.json`` /
``template_spec.json``.  Pure standard library, offline, Linux-safe.

Frozen interface (CONTRACT §4.1)::

    verify(work_dir, output_dir) -> dict
        reads  <work>/{case_manifest,template_spec,source_inventory,slides_plan}.json
               and the deck inside <output>
        writes <work>/fidelity_report.json + <work>/format_audit.json

``skipped`` is never ``pass``: every check that could not observe anything is listed
in ``unchecked`` and counted in ``vacuous_skips``.
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import sys
import unicodedata
import zipfile
from xml.etree import ElementTree as ET

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import common  # noqa: E402  (frozen shared layer, read-only)

A = common.NS["a"]
P = common.NS["p"]
R = common.NS["r"]

WORK_REQUIRED = ("case_manifest.json", "template_spec.json", "source_inventory.json",
                 "slides_plan.json")

# tolerance for non-proportional figure scaling
ASPECT_TOL = 0.02
# two text shapes are "overlapping" above this share of the smaller one's area
OVERLAP_TOL = 0.25
# a paragraph is "unsourced" when this share of its tokens is absent from allowed text
UNSOURCED_RATIO = 0.6
# plan text must land on the deck with at least this token coverage
PLAN_LAND_RATIO = 0.8

_BOILER = set("目录 结论 参考文献 致谢 摘要 引言 方法 结果 讨论 实验 背景 目标 总结 谢谢 "
              "图表 页 幻灯 报告 研究 分析 数据 模型 论文 作者 单位 年月".split())

_STOP = set("的 了 与 和 及 在 是 为 对 从 到 由 以 并 等 中 上 下 不 有 也 就 都 而 或 其 该 "
            "the a an of in on to for and or is are was were be by with as at from that this".split())

# literal markup that must never reach a rendered slide
_MARKUP_RX = (("markdown_bold", re.compile(r"\*\*|__")),
              ("markdown_link", re.compile(r"\[[^\]]+\]\([^)]+\)")),
              ("html_br", re.compile(r"<br\s*/?>", re.I)),
              ("html_entity", re.compile(r"&nbsp;|&amp;|&#\d+;", re.I)))


# ======================================================================================
# small helpers
# ======================================================================================

def _qn(ns, tag):
    return "{%s}%s" % (ns, tag)


def _now():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _read_json(path):
    data = common.read_json(path)
    return data if isinstance(data, dict) else None


def _norm(s):
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"\s+", "", s).lower()


def _tokens(s):
    """CJK letters + ASCII/digit words. Deterministic, no dictionary needed.

    The CJK class is **letters only** (Ext-A / unified ideographs / kana / Hangul).
    It deliberately excludes U+2E80–U+33FF (CJK radicals and CJK symbols & punctuation,
    e.g. ``。`` U+3002, ``、`` U+3001, ``「`` U+300C) and full-width forms U+FF00–U+FFEF:
    punctuation is not a content token, otherwise every correctly punctuated sentence
    would "lose" coverage against its evidence text.
    """
    s = unicodedata.normalize("NFKC", s or "")
    out = []
    for m in re.finditer(r"[A-Za-z][A-Za-z0-9_.\-]*|\d+(?:\.\d+)?%?|"
                         r"[\u3400-\u4DBF\u4E00-\u9FFF\u3040-\u30FF\uAC00-\uD7A3]", s):
        out.append(m.group(0).lower())
    return out


def _content_tokens(s):
    return [t for t in _tokens(s) if t not in _STOP]


def _char_ngrams(s, n=12):
    """Character n-grams of the whitespace-stripped text (long-span traceability)."""
    s = _norm(s)
    if len(s) < n:
        return set()
    return {s[i:i + n] for i in range(len(s) - n + 1)}


# "statistic-bearing" numeric tokens: decimals, percentages, years, 4+ digit integers and
# numbers carrying a unit.  Bare 1-3 digit integers are left out on purpose — they are
# overwhelmingly structural (list/section/page numbering) and would make every deck fail.
# Only the *numeric core* is used as the token: whether a unit is written adjacent to the
# number differs between a source paragraph and a slide, so adjacency must not decide.
_UNIT_ALT = (r"mg/l|kg|km|km/h|m/s|℃|°c|hz|ml|mm|cm|gb|mb|kb|"
             r"小时|分钟|秒|天|周|个月|年|轮|类|个|段|倍|个百分点|层|条")
_NUM_CORE = r"(?:\d+\.\d+|\d+(?:\.\d+)?%|(?:19|20)\d{2}|\d{4,})"
_STAT_RX = re.compile("(%s)(?:%s)?" % (_NUM_CORE, _UNIT_ALT))
_UNIT_RX = re.compile("\\d+(?:\\.\\d+)?(%s)" % _UNIT_ALT)


def _strip_percent(tok):
    return tok[:-1] if tok.endswith("%") else tok


def _stat_tokens(s):
    """Numeric cores of statistic-bearing tokens (percent sign folded in)."""
    return set(_strip_percent(m.group(1)) for m in _STAT_RX.finditer(_norm(s)))


def _unit_tokens(s):
    return set(m.group(1) for m in _UNIT_RX.finditer(_norm(s)))


def _num_tokens(s):
    return set(re.findall(r"\d+(?:\.\d+)?%?", unicodedata.normalize("NFKC", s or "")))


def _rel_target(base_part, target):
    """Resolve an OPC relationship target relative to the part that owns the rels."""
    base_dir = os.path.dirname(base_part)
    t = (target or "").replace("\\", "/")
    if t.startswith("/"):
        return t.lstrip("/")
    return os.path.normpath(os.path.join(base_dir, t)).replace("\\", "/")


def _parse_rels(parts, part):
    """Return [(rid, type, target, mode, resolved_part)] for a part's .rels."""
    d = os.path.dirname(part)
    rels = "_rels/" + os.path.basename(part) + ".rels"
    name = (d + "/" + rels) if d else rels
    raw = parts.get(name)
    if not raw:
        return []
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return []
    out = []
    for rel in root:
        rid = rel.get("Id")
        rtype = rel.get("Type") or ""
        target = rel.get("Target") or ""
        mode = rel.get("TargetMode") or "Internal"
        out.append((rid, rtype, target, mode, _rel_target(part, target) if mode == "Internal" else None))
    return out


def _text_of(elem):
    return "".join(t.text or "" for t in elem.iter(_qn(A, "t")))


def _paras_of(txbody):
    """[(text, [run_sizes_1_100pt], lnSpc_pct)] for each a:p of a txBody."""
    out = []
    for p in txbody.findall(_qn(A, "p")):
        runs = p.findall(_qn(A, "r"))
        txt = "".join((r.findtext(_qn(A, "t")) or "") for r in runs)
        sizes = []
        for r in runs:
            rpr = r.find(_qn(A, "rPr"))
            if rpr is not None and rpr.get("sz"):
                try:
                    sizes.append(int(rpr.get("sz")))
                except ValueError:
                    pass
        ln = None
        ppr = p.find(_qn(A, "pPr"))
        if ppr is not None:
            sp = ppr.find(_qn(A, "lnSpc"))
            if sp is not None:
                pct = sp.find(_qn(A, "spcPct"))
                if pct is not None and pct.get("val"):
                    try:
                        ln = int(pct.get("val")) / 100000.0
                    except ValueError:
                        ln = None
            if not sizes:
                dpr = ppr.find(_qn(A, "defRPr"))
                if dpr is not None and dpr.get("sz"):
                    try:
                        sizes.append(int(dpr.get("sz")))
                    except ValueError:
                        pass
        out.append({"text": txt, "sizes": sizes, "lnSpc": ln})
    return out


def _xfrm(shape):
    sp = shape.find(_qn(P, "spPr"))
    x = None
    if sp is not None:
        x = sp.find(_qn(A, "xfrm"))
    if x is None:
        return None
    off = x.find(_qn(A, "off"))
    ext = x.find(_qn(A, "ext"))
    if off is None or ext is None:
        return None
    try:
        return {"x": int(off.get("x", 0)), "y": int(off.get("y", 0)),
                "cx": int(ext.get("cx", 0)), "cy": int(ext.get("cy", 0))}
    except (TypeError, ValueError):
        return None


def _ph_of(shape):
    nv = shape.find(_qn(P, "nvSpPr"))
    if nv is None:
        return None
    pr = nv.find(_qn(P, "nvPr"))
    if pr is None:
        return None
    ph = pr.find(_qn(P, "ph"))
    if ph is None:
        return None
    return {"idx": ph.get("idx") or "0", "type": ph.get("type") or "body"}


def _ph_of_pic(shape):
    nv = shape.find(_qn(P, "nvPicPr"))
    if nv is None:
        return None
    pr = nv.find(_qn(P, "nvPr"))
    if pr is None:
        return None
    ph = pr.find(_qn(P, "ph"))
    if ph is None:
        return None
    return {"idx": ph.get("idx") or "0", "type": ph.get("type") or "pic"}


def _bodypr(elem):
    tx = elem.find(_qn(P, "txBody"))
    if tx is None:
        return None
    return tx.find(_qn(A, "bodyPr"))


def _insets(bodypr):
    def get(name, default):
        if bodypr is None:
            return default
        try:
            return int(bodypr.get(name, default))
        except (TypeError, ValueError):
            return default
    return (get("lIns", 91440), get("rIns", 91440), get("tIns", 45720), get("bIns", 45720))


def _autofit_scale(bodypr):
    if bodypr is None:
        return None
    naf = bodypr.find(_qn(A, "normAutofit"))
    if naf is None:
        return None
    val = naf.get("fontScale")
    if not val:
        return None
    try:
        return int(val) / 100000.0
    except ValueError:
        return None


_NAME_ROLE_RULES = (
    ("figurecaption", "caption"), ("tablecaption", "caption"), ("caption", "caption"),
    ("title", "title"), ("subtitle", "subtitle"),
    ("contentplaceholder", "body"), ("substitutedbody", "body"),
    ("textplaceholder", "body"), ("body", "body"),
)


def _layout_font_scale(layout_ph, slide_n, ph):
    """Inherited ``normAutofit fontScale`` from the slide's layout placeholder (1.0 if none)."""
    mapping = (layout_ph or {}).get(slide_n) or {}
    entry = mapping.get((ph or {}).get("idx"))
    if not entry:
        return 1.0
    body_pr = entry.get("body_pr") if isinstance(entry.get("body_pr"), dict) else {}
    for key in ("normAutofit_fontScale", "font_scale"):
        v = body_pr.get(key)
        if isinstance(v, float) and 0 < v <= 1:
            return v
        if isinstance(v, int):
            if 0 < v <= 1:
                return float(v)
            if v > 1000:
                return v / 100000.0
    return 1.0


def _role_and_source(ph, name=None):
    """Role of a shape: ``<p:ph>`` first, else the agreed shape-name convention.

    The engine names substituted shapes deterministically, so a text box without a
    placeholder is identified structurally ("Figure Caption" -> caption) instead of by
    guessing from its text.  Returns ``(role, source)``.
    """
    if ph:
        return _role_of(ph), "ph:" + str(ph.get("type"))
    if name:
        key = re.sub(r"[^a-z]", "", name.lower())
        for pat, role in _NAME_ROLE_RULES:
            if key == pat or key.startswith(pat) or key.endswith(pat):
                return role, "name:%s" % name
    return "body", "default"


def _role_of(ph, name=None):
    if not ph:
        return _role_and_source(None, name)[0]
    t = ph.get("type", "body")
    if t in ("title", "ctrTitle"):
        return "title"
    if t == "subTitle":
        return "subtitle"
    if t in ("sldNum", "ftr", "dt"):
        return "meta"
    if t == "pic":
        return "pic"
    return "body"


# ======================================================================================
# template / source model
# ======================================================================================

def _theme_of(parts, theme_part):
    raw = parts.get(theme_part)
    if not raw:
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    out = {"fonts": {}, "colors": {}, "colors_alt": {}}
    fs = root.find(".//" + _qn(A, "fontScheme"))
    if fs is not None:
        for kind in ("majorFont", "minorFont"):
            node = fs.find(_qn(A, kind))
            if node is None:
                continue
            for sub in ("latin", "ea", "cs"):
                el = node.find(_qn(A, sub))
                out["fonts"]["%s_%s" % (kind[:-4], sub)] = (el.get("typeface") if el is not None else None)
    cs = root.find(".//" + _qn(A, "clrScheme"))
    if cs is not None:
        for child in cs:
            tag = common.localname(child.tag)
            v, alt = None, None
            for sub in ("srgbClr", "sysClr", "schemeClr"):
                el = child.find(_qn(A, sub))
                if el is not None:
                    v = el.get("val")
                    # system colours carry both a symbolic name and the resolved sRGB hex;
                    # template_spec may legitimately record either one
                    if sub == "sysClr":
                        alt = el.get("lastClr")
                    break
            out["colors"][tag] = v
            if alt:
                out["colors_alt"][tag] = [alt]
    return out


def _layout_ph_map(parts, layout_part):
    """{idx: {'type':..., 'xfrm':..., 'sz':..., 'name':...}} for one layout part."""
    raw = parts.get(layout_part)
    if not raw:
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    out = {}
    for sp in root.iter(_qn(P, "sp")):
        ph = _ph_of(sp)
        if not ph:
            continue
        entry = {"type": ph["type"], "xfrm": _xfrm(sp), "sz": None, "name": None}
        nv = sp.find(_qn(P, "nvSpPr"))
        if nv is not None:
            c = nv.find(_qn(P, "cNvPr"))
            if c is not None:
                entry["name"] = c.get("name")
        tx = sp.find(_qn(P, "txBody"))
        if tx is not None:
            lst = tx.find(_qn(A, "lstStyle"))
            if lst is not None:
                l1 = lst.find(_qn(A, "lvl1pPr"))
                if l1 is not None:
                    d = l1.find(_qn(A, "defRPr"))
                    if d is not None and d.get("sz"):
                        try:
                            entry["sz"] = int(d.get("sz"))
                        except ValueError:
                            pass
            bp = tx.find(_qn(A, "bodyPr"))
            if entry["sz"] is None and bp is not None:
                d = bp.find(_qn(A, "defRPr"))
                if d is not None and d.get("sz"):
                    try:
                        entry["sz"] = int(d.get("sz"))
                    except ValueError:
                        pass
        out[ph["idx"]] = entry
    return out


def _master_ph_map(parts, master_part):
    m = _layout_ph_map(parts, master_part)
    return m or {}


def _slide_parts(parts):
    names = [n for n in parts if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
    return sorted(names, key=lambda n: int(re.search(r"(\d+)", os.path.basename(n)).group(1)))


def _deck_theme_parts(parts):
    return sorted(n for n in parts if re.match(r"^ppt/theme/theme\d+\.xml$", n))


def _midia(parts):
    return [n for n in parts if n.startswith("ppt/media/")]


# ======================================================================================
# deck extraction
# ======================================================================================

def _extract_deck(parts, spec):
    """Structured deck model computed only from the pptx bytes."""
    slides = []
    for sp_name in _slide_parts(parts):
        raw = parts[sp_name]
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            slides.append({"part": sp_name, "error": "unparsable xml", "shapes": []})
            continue
        rels = _parse_rels(parts, sp_name)
        layout = None
        notes = None
        for _rid, rtype, _tgt, mode, resolved in rels:
            if mode != "Internal" or not resolved:
                continue
            if rtype.endswith("/slideLayout"):
                layout = resolved
            elif rtype.endswith("/notesSlide"):
                notes = resolved
        shapes = []
        pics = []
        for sp in root.iter(_qn(P, "sp")):
            ph = _ph_of(sp)
            tx = sp.find(_qn(P, "txBody"))
            paras = _paras_of(tx) if tx is not None else []
            text = "\n".join(p["text"] for p in paras)
            nv = sp.find(_qn(P, "nvSpPr"))
            name = None
            if nv is not None:
                c = nv.find(_qn(P, "cNvPr"))
                if c is not None:
                    name = c.get("name")
            role, role_src = _role_and_source(ph, name)
            shapes.append({"kind": "sp", "name": name, "ph": ph,
                           "role": role, "role_source": role_src, "xfrm": _xfrm(sp),
                           "paras": paras, "text": text,
                           "autofit": _autofit_scale(_bodypr(sp)),
                           "txbody": tx})
        for pic in root.iter(_qn(P, "pic")):
            nv = pic.find(_qn(P, "nvPicPr"))
            c = nv.find(_qn(P, "cNvPr")) if nv is not None else None
            name = c.get("name") if c is not None else None
            ph = _ph_of_pic(pic)
            blip = pic.find(".//" + _qn(A, "blip"))
            embed = blip.get(_qn(R, "embed")) if blip is not None else None
            target = None
            for rid, rtype, _tgt, mode, resolved in rels:
                if rid == embed and mode == "Internal":
                    target = resolved
                    break
            pics.append({"kind": "pic", "name": name, "ph": ph, "role": "pic",
                         "role_source": "pic", "xfrm": _xfrm(pic), "target": target,
                         "paras": [], "text": "", "autofit": None, "txbody": None})
            shapes.append(pics[-1])
        notes_text = ""
        if notes and notes in parts:
            try:
                nroot = ET.fromstring(parts[notes])
                notes_text = "\n".join(t.text or "" for t in nroot.iter(_qn(A, "t")))
            except ET.ParseError:
                notes_text = ""
        # editable tables: <p:graphicFrame><a:graphic><a:graphicData><a:tbl>
        tables = []
        for tbl in root.iter(_qn(A, "tbl")):
            cells = []
            for tr in tbl.findall(_qn(A, "tr")):
                row = []
                for tc in tr.findall(_qn(A, "tc")):
                    row.append("\n".join(t.text or "" for t in tc.iter(_qn(A, "t"))))
                cells.append(row)
            tables.append({"cells": cells, "rows": len(cells),
                           "cols": max([len(r) for r in cells] or [0])})
        slides.append({"part": sp_name, "n": len(slides) + 1, "layout": layout,
                       "notes_part": notes, "notes_text": notes_text,
                       "shapes": shapes, "pics": pics, "tables": tables, "root": root,
                       "slide_num_fld": any(True for f in root.iter(_qn(A, "fld"))
                                            if f.get("type") in ("slidenum",)),
                       "footer_fld": any(True for f in root.iter(_qn(A, "fld"))
                                         if f.get("type") in ("ftr",))})
    return slides


# ======================================================================================
# fidelity
# ======================================================================================

def _source_corpus(inv):
    chunks = []
    for u in inv.get("units") or []:
        if isinstance(u, dict):
            chunks.append(u.get("text") or "")
    for key in ("figures", "tables"):
        for it in inv.get(key) or []:
            if isinstance(it, dict):
                chunks.append(it.get("caption") or "")
                chunks.append(it.get("label") or "")
    for s in inv.get("sections") or []:
        if isinstance(s, dict):
            chunks.append(s.get("title") or "")
    return "\n".join(chunks)


def _plan_corpus(plan):
    chunks = []
    for sl in plan.get("slides") or []:
        if not isinstance(sl, dict):
            continue
        chunks.append(sl.get("title") or "")
        chunks.append(sl.get("subtitle") or "")
        chunks.append(sl.get("notes") or "")
        for b in sl.get("bullets") or []:
            if isinstance(b, dict):
                chunks.append(b.get("text") or "")
        for f in sl.get("figures") or []:
            if isinstance(f, dict):
                chunks.append(f.get("caption") or "")
        for c in sl.get("claims") or []:
            chunks.append(c if isinstance(c, str) else "")
    return "\n".join(chunks)


def _plan_content_corpus(plan):
    """Plan *content* only: speaker notes are author-written narration.

    Notes are governed by the dedicated ``notes_sourced`` check; including them here as
    well would double-count them and contradict that check.
    """
    chunks = []
    for sl in plan.get("slides") or []:
        if not isinstance(sl, dict):
            continue
        chunks.append(sl.get("title") or "")
        chunks.append(sl.get("subtitle") or "")
        for b in sl.get("bullets") or []:
            if isinstance(b, dict):
                chunks.append(b.get("text") or "")
        for f in sl.get("figures") or []:
            if isinstance(f, dict):
                chunks.append(f.get("caption") or "")
        for c in sl.get("claims") or []:
            chunks.append(c if isinstance(c, str) else "")
    return "\n".join(chunks)


def _plan_slide_texts(plan):
    out = []
    for i, sl in enumerate(plan.get("slides") or [], 1):
        if not isinstance(sl, dict):
            out.append("")
            continue
        bits = [sl.get("title") or "", sl.get("subtitle") or ""]
        for b in sl.get("bullets") or []:
            if isinstance(b, dict):
                bits.append(b.get("text") or "")
        bits.append(sl.get("notes") or "")
        out.append("\n".join(bits))
    return out


def _unit_ids_of_plan(plan):
    ids = set()
    for sl in plan.get("slides") or []:
        if not isinstance(sl, dict):
            continue
        for b in sl.get("bullets") or []:
            if isinstance(b, dict):
                for e in b.get("evidence") or []:
                    if isinstance(e, str):
                        ids.add(e)
        for e in (sl.get("evidence") or []) + (sl.get("notes_evidence") or []):
            if isinstance(e, str):
                ids.add(e)
    return ids


def _check(name, status, expected, observed, detail="", evidence=None):
    return {"name": name, "status": status, "expected": expected, "observed": observed,
            "detail": detail, "evidence": evidence or []}


# ======================================================================================
# v1.3 capability audit helpers (ADDENDUM v1.3 §A5)
#
# The four new checks are judged only against the *real bytes* plus the frozen v1.3
# intermediates.  Applicability is explicit: a check that does not apply to this run is
# recorded with status ``not_applicable`` (never ``pass``); a check that should apply but
# has no evidence is a ``fail`` -- silently skipping is exactly the failure mode v1.2
# closed.
# ======================================================================================

V13_CHECKS = ("start_deck_preserved", "table_rendered", "constraints_honoured",
              "notes_evidence_present")

# per start-deck text unit: share of its content tokens that must survive onto the deck
START_DECK_UNIT_RATIO = 0.8
# speaker-notes content must be this well covered by the notes_evidence source texts.
# ``notes`` is an authored script, not a quote: after removing the connector/stopword set
# below, the remaining *content* tokens must be >= 90% present in the evidence texts and
# every statistic must be present verbatim.
NOTES_EVIDENCE_RATIO = 0.9
# narration/connector vocabulary that carries no fact (ADDENDUM §A4 allows 虚词与连接词)
_NOTES_STOP = _STOP | set(
    "本 页 要 点 接 下 来 看 先 其 次 然 后 我 们 这 里 那 将 会 说 明 及 如 所 示 请 注 意 "
    "问 题 一 二 三 四 五 六 七 八 九 十 第 个 部 分 内 容 讲 解 过 渡 张 幻 灯 片 页 面 "
    "首 先 同 时 因 此 所 以 但 是 而 且 另 外 综 上 以 下 以 上 具 体 情 况 相 关 方 面 "
    "then next now here this that thus also however therefore finally first second".split())


def _plan_is_v13(plan):
    """A v1.3 plan always declares ``mode``; v1.2 plans are not judged on v1.3 features."""
    return (isinstance(plan, dict) and isinstance(plan.get("mode"), str)
            and plan["mode"].strip() != "")


def _load_json(work, name):
    data = common.read_json(os.path.join(work, name))
    return data if isinstance(data, dict) else {}


def _read_case_json(work, manifest):
    """Locate the real frozen ``case.json`` and return ``(case_dict, path)``."""
    cands = []
    for entry in (manifest.get("inputs") or []):
        if not isinstance(entry, dict):
            continue
        rel = (entry.get("rel") or "").replace("\\", "/")
        if rel.endswith("case.json"):
            if entry.get("abs"):
                cands.append(entry["abs"])
            cands.append(os.path.join(work, "inputs", rel))
    cands += [os.path.join(work, "inputs", "case.json"),
              os.path.join(work, "case.json")]
    for cand in cands:
        try:
            if cand and os.path.isfile(cand):
                with open(cand, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict):
                    return data, cand
        except (OSError, ValueError):
            continue
    return {}, None


def _case_constraints(work, manifest):
    case, path = _read_case_json(work, manifest)
    cons = case.get("constraints") if isinstance(case, dict) else None
    return (cons if isinstance(cons, dict) else {}), path


def _start_deck_spec(work):
    return _load_json(work, "start_deck_spec.json")


def _adaptation_actions(work):
    acts = (_load_json(work, "adaptation_actions.json") or {}).get("actions") or []
    return [a for a in acts if isinstance(a, dict)]


def _slide_index_of(action):
    for key in ("slide", "start_slide", "index", "n"):
        try:
            return int(action.get(key))
        except (TypeError, ValueError):
            continue
    return None


def _dropped_start_slides(actions):
    """{start-deck page index: drop reason} for recorded ``start_deck_slide_dropped``."""
    out = {}
    for a in actions:
        if a.get("kind") != "start_deck_slide_dropped":
            continue
        idx = _slide_index_of(a)
        if idx is None:
            continue
        reason = a.get("reason") or a.get("why") or a.get("note") or ""
        out[idx] = str(reason).strip()
    return out


def _preserved_start_slides(actions):
    out = set()
    for a in actions:
        if a.get("kind") != "start_deck_slide_preserved":
            continue
        idx = _slide_index_of(a)
        if idx is not None:
            out.add(idx)
    return out


def _norm_cell(value):
    s = unicodedata.normalize("NFKC", "" if value is None else str(value))
    return re.sub(r"\s+", " ", s).strip()


def _table_matches(cells, want):
    if len(cells) != len(want):
        return False
    for got_row, want_row in zip(cells, want):
        if len(got_row) != len(want_row):
            return False
        for got, exp in zip(got_row, want_row):
            if _norm_cell(got) != _norm_cell(exp):
                return False
    return True


def _inventory_item_text(item):
    """Fullest available text of a source_inventory figure/table entry.

    Order: ``label`` + ``caption_plain`` (markdown stripped; falls back to ``caption``),
    then every real ``header``/``rows`` cell.  Table numbers live in the cells, so they
    must be part of the evidence corpus or a script citing the table looks unsourced.
    """
    if not isinstance(item, dict):
        return ""
    parts = [str(item.get("label") or "")]
    plain = item.get("caption_plain")
    if isinstance(plain, str) and plain.strip():
        parts.append(plain)
    else:
        parts.append(str(item.get("caption") or ""))
    for key in ("header", "rows"):
        for row in (item.get(key) or []):
            if isinstance(row, (list, tuple)):
                parts.extend(str(cell) for cell in row)
            else:
                parts.append(str(row))
    return " ".join(p for p in parts if p)


def _evidence_index(inv, start_spec=None):
    """notes_evidence id -> source text (units / sections / figures / tables).

    ADDENDUM v1.3 §A4 also allows ``start_deck:<part>`` ids: the starting deck is
    organiser-supplied draft material, so a script about it is evidenced by the starting
    deck page itself (``<work>/start_deck_spec.json`` ``slide_details[].text_units``).
    """
    idx = {}
    for it in (inv.get("units") or []):
        if isinstance(it, dict) and it.get("id"):
            idx[str(it["id"])] = it.get("text") or ""
    for it in (inv.get("sections") or []):
        if isinstance(it, dict) and it.get("id"):
            idx[str(it["id"])] = it.get("title") or ""
    for key in ("figures", "tables"):
        for it in (inv.get(key) or []):
            if isinstance(it, dict) and it.get("id"):
                idx[str(it["id"])] = _inventory_item_text(it)
    for sd in ((start_spec or {}).get("slide_details") or []):
        if not isinstance(sd, dict):
            continue
        part = sd.get("part")
        if not part:
            continue
        texts = [u["text"] for u in (sd.get("text_units") or [])
                 if isinstance(u, dict) and (u.get("text") or "").strip()]
        texts = ([sd.get("title")] if (sd.get("title") or "").strip() else []) + texts
        idx["start_deck:" + str(part)] = "\n".join(dict.fromkeys(texts))
    return idx


def _notes_content_tokens(text):
    """Content tokens for speaker notes: word chars minus the connector/stopword set."""
    return [t for t in _tokens(text or "") if t not in _NOTES_STOP]


def _start_deck_texts(spec):
    """Every text unit in ``start_deck_spec.slide_details`` (v1.3 schema)."""
    out = []
    for sd in (spec.get("slide_details") or []):
        if not isinstance(sd, dict):
            continue
        for u in (sd.get("text_units") or []):
            if isinstance(u, dict) and (u.get("text") or "").strip():
                out.append(u["text"])
    return out


# --------------------------------------------------------------------------------------
# v1.3 fidelity checks
# --------------------------------------------------------------------------------------

def _fidelity_start_deck(deck, plan, work):
    """start_deck_preserved: every starting-deck text unit survives, or its page is waived."""
    name = "start_deck_preserved"
    if not _plan_is_v13(plan):
        return None, [], [], {"name": name,
                              "reason": "plan has no v1.3 `mode` key (legacy schema): "
                                        "start-deck adaptation cannot be judged"}
    if plan.get("mode") != "adapt_start_deck":
        return None, [], [], {"name": name,
                              "reason": "plan.mode=%r: this run does not adapt a starting deck"
                                        % plan.get("mode")}
    spec = _start_deck_spec(work)
    details = [d for d in (spec.get("slide_details") or []) if isinstance(d, dict)]
    if not spec or spec.get("ok") is not True:
        chk = _check(name, "fail", "readable start_deck_spec.json with ok=true", 0,
                     "start_deck_spec.json missing or ok!=true: %r"
                     % ((spec.get("notes") if isinstance(spec.get("notes"), list) else
                         spec.get("notes")) or "absent"))
        return chk, [{"kind": "start_deck_unauditable", "slide": None, "text": "",
                      "why": "no readable start_deck_spec.json"}], [], None
    if not details:
        chk = _check(name, "fail", "v1.3 slide_details in start_deck_spec.json", 0,
                     "start_deck_spec.json has no slide_details (v1.2 minimal schema): "
                     "per-page text preservation cannot be audited")
        return chk, [{"kind": "start_deck_unauditable", "slide": None, "text": "",
                      "why": "start_deck_spec.json lacks slide_details"}], [], None

    deck_text = "\n".join(sh.get("text") or "" for sl in deck
                          for sh in sl["shapes"] if sh.get("kind") == "sp")
    deck_tokens = set(_content_tokens(deck_text))
    actions = _adaptation_actions(work)
    dropped = _dropped_start_slides(actions)
    preserved_actions = _preserved_start_slides(actions)
    plan_source_pages = {int(sl.get("source_slide")) for sl in (plan.get("slides") or [])
                         if isinstance(sl, dict) and str(sl.get("source_slide", "")).strip().isdigit()}
    checked, lost, waived, informational = 0, [], [], []
    pages_with_text = 0
    for sd in details:
        idx = sd.get("index")
        units = [u for u in (sd.get("text_units") or [])
                 if isinstance(u, dict) and (u.get("text") or "").strip()]
        if not units:
            continue
        pages_with_text += 1
        reason = dropped.get(idx)
        if reason:
            waived.append({"slide": idx, "reason": reason})
            if idx not in plan_source_pages and idx not in preserved_actions:
                informational.append({"kind": "start_slide_drop_not_in_plan", "slide": idx,
                                      "why": "page dropped with a reason but no plan/action trace "
                                             "of it being considered"})
            continue
        if idx in dropped:
            informational.append({"kind": "start_slide_drop_without_reason", "slide": idx,
                                  "why": "start_deck_slide_dropped recorded without a reason; "
                                         "not a valid waiver"})
        if idx not in plan_source_pages and idx not in preserved_actions:
            informational.append({"kind": "start_slide_unreferenced", "slide": idx,
                                  "why": "page not mapped by any plan slide and no "
                                         "start_deck_slide_preserved action recorded"})
        for u in units:
            toks = _content_tokens(u.get("text") or "")
            if not toks:
                continue
            checked += 1
            cov = len([t for t in toks if t in deck_tokens]) / float(len(toks))
            if cov < START_DECK_UNIT_RATIO:
                lost.append({"slide": idx, "where": u.get("where"),
                             "text": (u.get("text") or "")[:120],
                             "coverage": round(cov, 3),
                             "why": "starting-deck text unit missing from the produced deck"})

    if not pages_with_text:
        return (_check(name, "skipped", "starting-deck text units", 0,
                       "start_deck_spec slide_details carry no text units"),
                [], [], None)
    if not lost:
        detail = ("pages with text=%d units checked=%d waived=%d"
                  % (pages_with_text, checked, len(waived)))
        chk = _check(name, "pass", "every starting-deck text unit survives (or its page is "
                     "dropped with a recorded reason)", {"lost": 0, "waived": len(waived)}, detail)
        return chk, [], informational, None
    chk = _check(name, "fail", "every starting-deck text unit survives (or its page is "
                 "dropped with a recorded reason)", len(lost),
                 "pages with text=%d units checked=%d lost=%d waived=%d"
                 % (pages_with_text, checked, len(lost), len(waived)))
    chk["detail_extra"] = lost[:10]
    return chk, [{"kind": "start_deck_text_dropped", "slide": l["slide"],
                  "text": l["text"], "why": l["why"]} for l in lost], informational, None


def _fidelity_table(deck, plan):
    """table_rendered: every planned table exists as a real <a:tbl> with matching cells."""
    name = "table_rendered"
    planned = []
    for i, sl in enumerate(plan.get("slides") or [], 1):
        if not isinstance(sl, dict):
            continue
        for t in (sl.get("tables") or []):
            if isinstance(t, dict):
                planned.append((i, t))
    if not planned:
        return None, [], [], {"name": name, "reason": "slides_plan declares no tables"}
    deck_tables = []
    for sl in deck:
        for tb in (sl.get("tables") or []):
            deck_tables.append((sl["n"], tb))

    lost = []
    for i, t in planned:
        header = [_norm_cell(x) for x in (t.get("header") or [])]
        rows = [[_norm_cell(c) for c in r] for r in (t.get("rows") or [])
                if isinstance(r, (list, tuple))]
        expected = ([header] if any(header) else []) + rows
        exp_r = len(expected)
        exp_c = max([len(r) for r in expected] or [0])
        want = [r + [""] * (exp_c - len(r)) for r in expected]
        match = None
        # prefer the plan slide index, then accept the table anywhere (split pages)
        for n, tb in deck_tables:
            if n == i and _table_matches(tb.get("cells") or [], want):
                match = (n, tb)
                break
        if match is None:
            for n, tb in deck_tables:
                if _table_matches(tb.get("cells") or [], want):
                    match = (n, tb)
                    break
        if match is not None:
            continue
        observed = [{"slide": n, "rows": tb.get("rows"), "cols": tb.get("cols")}
                    for n, tb in deck_tables]
        why = "no <a:tbl> carries the planned %dx%d table" % (exp_r, exp_c)
        if not deck_tables:
            slide_text = "\n".join(sh.get("text") or "" for sl in deck if sl["n"] == i
                                   for sh in sl["shapes"] if sh.get("kind") == "sp")
            flat = [c for r in want for c in r if c]
            hit = sum(1 for c in flat if c and c in _norm_cell(slide_text))
            if flat and hit / float(len(flat)) >= 0.5:
                why = ("planned table was flattened into plain text "
                       "(%d/%d cell strings found in text shapes)" % (hit, len(flat)))
            elif any(sh.get("kind") == "pic" for sl in deck if sl["n"] == i
                     for sh in sl["shapes"]):
                why = "planned table was replaced by a picture"
        lost.append({"slide": i, "caption": (t.get("caption") or "")[:80],
                     "expected": "%dx%d" % (exp_r, exp_c),
                     "observed": observed, "why": why})
    if not lost:
        chk = _check(name, "pass", "every planned table is a real <a:tbl> with matching size "
                     "and cell text", 0,
                     "planned tables=%d rendered=%d" % (len(planned), len(planned)))
        return chk, [], [], None
    chk = _check(name, "fail", "every planned table is a real <a:tbl> with matching size "
                 "and cell text", len(lost),
                 "planned tables=%d unrendered=%d deck_tables=%d"
                 % (len(planned), len(lost), len(deck_tables)))
    chk["detail_extra"] = lost[:10]
    return chk, [{"kind": "table_not_rendered", "slide": l["slide"], "text": l["caption"],
                  "why": l["why"]} for l in lost], [], None


def _fidelity_notes(deck, plan, inv, start_spec=None):
    """notes_evidence_present: non-empty, resolvable, sourced speaker notes.

    Rules (published in ``detail`` for review):
      * content tokens = word tokens minus the connector/stopword set (``_NOTES_STOP``);
      * at least ``NOTES_EVIDENCE_RATIO`` (90%) of them must occur in the evidence texts;
      * every statistic in the script must occur in the evidence texts verbatim;
      * ``unresolved`` (unknown/absent evidence id) and ``low_overlap`` (sourced but not
        enough overlap) are counted separately, never merged into one silent skip.
    """
    name = "notes_evidence_present"
    if not _plan_is_v13(plan):
        return None, [], [], {"name": name,
                              "reason": "plan has no v1.3 `mode` key (legacy schema): "
                                        "notes_evidence is not required there"}
    slides = [sl for sl in (plan.get("slides") or []) if isinstance(sl, dict)]
    with_notes = [(i, sl) for i, sl in enumerate(slides, 1)
                  if (sl.get("notes") or "").strip()]
    if not with_notes:
        return None, [], [], {"name": name, "reason": "plan carries no speaker notes"}

    index = _evidence_index(inv, start_spec)
    threshold = NOTES_EVIDENCE_RATIO
    bad, corpora = [], []
    unresolved = low_overlap = stat_leak = missing_evidence = 0
    for i, sl in with_notes:
        ev = sl.get("notes_evidence")
        ev = [e for e in ev if isinstance(e, str) and e.strip()] if isinstance(ev, list) else []
        if not ev:
            missing_evidence += 1
            bad.append({"slide": i, "kind": "missing_evidence",
                        "why": "notes_evidence is missing or empty"})
            continue
        unknown = [e for e in ev if e not in index]
        if unknown:
            unresolved += 1
            bad.append({"slide": i, "kind": "unresolved",
                        "why": "unresolved notes_evidence ids %r" % unknown})
            continue
        corpus = "\n".join(index[e] for e in ev)
        corpora.append(corpus)
        note = sl.get("notes") or ""
        toks = _notes_content_tokens(note)
        if toks:
            ctoks = set(_notes_content_tokens(corpus))
            cov = len([t for t in toks if t in ctoks]) / float(len(toks))
            if cov < threshold:
                low_overlap += 1
                bad.append({"slide": i, "kind": "low_overlap",
                            "why": "only %.0f%% of notes content tokens occur in the "
                                   "evidence texts (threshold %.0f%%)"
                                   % (cov * 100.0, threshold * 100.0)})
        stat_bad = sorted(_stat_tokens(note) - _stat_tokens(corpus))
        if stat_bad:
            stat_leak += 1
            bad.append({"slide": i, "kind": "unsourced_number",
                        "why": "notes numbers absent from evidence: %s" % stat_bad})

    deck_notes = "\n".join((sl.get("notes_text") or "") for sl in deck
                           if (sl.get("notes_text") or "").strip())
    plan_notes = "\n".join((sl.get("notes") or "") for _i, sl in with_notes)
    deck_low_overlap = deck_stat_leak = 0
    if not deck_notes.strip():
        bad.append({"slide": None, "kind": "missing_delivery",
                    "why": "plan has speaker notes but the pptx carries no notes text"})
    else:
        allowed = set(_notes_content_tokens("\n".join(corpora) + "\n" + plan_notes)) | _BOILER
        dtoks = _notes_content_tokens(deck_notes)
        if dtoks:
            cov = len([t for t in dtoks if t in allowed]) / float(len(dtoks))
            if cov < threshold:
                deck_low_overlap = 1
                bad.append({"slide": None, "kind": "deck_low_overlap",
                            "why": "deck notes text is not backed by any notes_evidence "
                                   "(content coverage %.0f%%, threshold %.0f%%)"
                                   % (cov * 100.0, threshold * 100.0)})
        stat_bad = sorted(_stat_tokens(deck_notes) - _stat_tokens("\n".join(corpora)))
        if stat_bad:
            deck_stat_leak = 1
            bad.append({"slide": None, "kind": "deck_unsourced_number",
                        "why": "deck notes numbers absent from evidence: %s" % stat_bad})

    rules = ("rules: content tokens = CJK letters + numbers/ASCII words (punctuation is "
             "never a content token) minus %d connector/stopword tokens; content-token "
             "coverage >= %.0f%%; every statistic must appear verbatim in the evidence "
             "texts; unresolved=%d low_overlap=%d unsourced_number=%d missing_evidence=%d "
             "deck_low_overlap=%d deck_unsourced_number=%d"
             % (len(_NOTES_STOP), threshold * 100.0, unresolved, low_overlap,
                stat_leak, missing_evidence, deck_low_overlap, deck_stat_leak))
    if not bad:
        chk = _check(name, "pass", "every slide with notes has non-empty resolvable "
                     "notes_evidence and a script traceable to it", 0,
                     "slides with notes=%d; %s" % (len(with_notes), rules))
        return chk, [], [], None
    chk = _check(name, "fail", "every slide with notes has non-empty resolvable "
                 "notes_evidence and a script traceable to it", len(bad),
                 "slides with notes=%d violations=%d; %s"
                 % (len(with_notes), len(bad), rules))
    chk["detail_extra"] = bad[:10]
    return chk, [{"kind": "notes_unsourced", "slide": b.get("slide"), "text": "",
                  "why": b["why"]} for b in bad], [], None


def run_fidelity(work, parts, deck, manifest, inv, plan):
    checks = []
    violations = []
    informational = []
    not_applicable = []
    metrics = {"deck_text_units": 0, "deck_units_traceable": 0,
               "source_units_covered": 0, "source_units_total": 0,
               "unsourced_deck_strings": [], "figure_refs_checked": 0, "figure_refs_ok": 0}

    src_corpus = _source_corpus(inv)
    # v1.3: a starting deck is real input material, so its own text is a legitimate
    # source for the adapted deck (otherwise every preserved sentence looks invented).
    start_spec = _start_deck_spec(work)
    start_text_chunks = _start_deck_texts(start_spec)
    if start_text_chunks:
        src_corpus = src_corpus + "\n" + "\n".join(start_text_chunks)
    plan_corpus = _plan_corpus(plan)
    allowed = set(_content_tokens(src_corpus)) | set(_content_tokens(plan_corpus))
    allowed |= _BOILER
    src_tokens = set(_content_tokens(src_corpus))
    # long-span traceability: a rewritten sentence keeps words but not 12-char spans
    src_ngrams = _char_ngrams(src_corpus) | _char_ngrams(plan_corpus)

    # --- no_invented_text: every deck paragraph must be traceable to source/plan ---
    unsourced = []
    deck_paras = []
    for sl in deck:
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            for para in sh["paras"]:
                txt = (para["text"] or "").strip()
                if txt:
                    deck_paras.append((sl["n"], sh.get("role"), txt))
    traceable = 0
    for n, role, txt in deck_paras:
        toks = _content_tokens(txt)
        if not toks:
            traceable += 1
            continue
        bad = [t for t in toks if t not in allowed]
        ratio = len(bad) / float(len(toks))
        ng = _char_ngrams(txt)
        ng_frac = (len(ng & src_ngrams) / float(len(ng))) if ng else 1.0
        invented = (len(_norm(txt)) >= 24 and ng and ng_frac < 0.35) or \
                   (len(toks) >= 4 and ratio >= UNSOURCED_RATIO)
        if invented:
            unsourced.append({"slide": n, "role": role, "text": txt[:160],
                              "why": "token-absent %.0f%%, 12-gram match %.0f%%"
                                     % (ratio * 100, ng_frac * 100)})
        else:
            traceable += 1
    metrics["deck_text_units"] = len(deck_paras)
    metrics["deck_units_traceable"] = traceable
    metrics["unsourced_deck_strings"] = [u["text"] for u in unsourced]
    if not deck_paras:
        checks.append(_check("no_invented_text", "skipped", ">=1 deck paragraph",
                            0, "deck has no text paragraphs"))
    else:
        checks.append(_check("no_invented_text", "fail" if unsourced else "pass",
                             "0 unsourced deck paragraphs", len(unsourced),
                             "unsourced=%d/%d" % (len(unsourced), len(deck_paras))))
    for u in unsourced:
        violations.append({"kind": "unsourced_text", "slide": u["slide"],
                           "text": u["text"], "why": u["why"]})

    # --- plan_traceable: plan *content* must not invent facts (notes excluded) ---
    plan_paras = [t for t in _plan_content_corpus(plan).split("\n") if t.strip()]
    plan_bad = []
    for txt in plan_paras:
        toks = _content_tokens(txt)
        if len(toks) < 4:
            continue
        bad = [t for t in toks if t not in (src_tokens | _BOILER)]
        if bad and len(bad) / float(len(toks)) >= UNSOURCED_RATIO:
            plan_bad.append(txt[:160])
    if not plan_paras:
        checks.append(_check("plan_traceable", "skipped", "plan content present", 0,
                             "slides_plan has no content text (notes are checked by notes_sourced)"))
    else:
        checks.append(_check("plan_traceable", "fail" if plan_bad else "pass",
                             "0 unsourced plan content paragraphs", len(plan_bad),
                             "unsourced content paragraphs=%d (notes excluded)"
                             % len(plan_bad)))
        for txt in plan_bad:
            violations.append({"kind": "unsourced_plan_text", "slide": None,
                               "text": txt, "why": "plan text absent from source"})

    # --- plan_content_landed: plan bullets/notes must actually appear on the deck ---
    deck_by_n = {}
    for sl in deck:
        deck_by_n[sl["n"]] = "\n".join(sh["text"] for sh in sl["shapes"] if sh["kind"] == "sp") \
            + "\n" + (sl.get("notes_text") or "")
    plan_slides = _plan_slide_texts(plan)
    landed_fail = []
    checked = 0
    for i, ptext in enumerate(plan_slides, 1):
        toks = _content_tokens(ptext)
        if not toks:
            continue
        checked += 1
        dtoks = set(_content_tokens(deck_by_n.get(i, "")))
        cov = len([t for t in toks if t in dtoks]) / float(len(toks))
        if cov < PLAN_LAND_RATIO:
            landed_fail.append({"slide": i, "coverage": round(cov, 3)})
    if not checked:
        checks.append(_check("plan_content_landed", "skipped", "plan content lands", 0,
                             "no plan slides with text"))
    else:
        checks.append(_check("plan_content_landed", "fail" if landed_fail else "pass",
                             "each plan slide coverage >= %.2f" % PLAN_LAND_RATIO,
                             len(landed_fail),
                             "slides below threshold=%d" % len(landed_fail)))
        for f in landed_fail:
            violations.append({"kind": "dropped_content", "slide": f["slide"],
                               "text": "", "why": "plan coverage %.2f" % f["coverage"]})

    # --- source_token_recall over plan-referenced units ---
    units = [u for u in (inv.get("units") or []) if isinstance(u, dict) and u.get("id")]
    referenced = _unit_ids_of_plan(plan)
    sel = [u for u in units if (u.get("id") in referenced)] if referenced else units
    metrics["source_units_total"] = len(sel)
    deck_all = set(_content_tokens("\n".join(deck_by_n.values())))
    covered = 0
    for u in sel:
        toks = _content_tokens(u.get("text") or "")
        if not toks:
            covered += 1
            continue
        cov = len([t for t in toks if t in deck_all]) / float(len(toks))
        if cov >= 0.5:
            covered += 1
    metrics["source_units_covered"] = covered
    if not sel:
        checks.append(_check("source_token_recall", "skipped", "source units", 0,
                             "source_inventory has no units"))
    else:
        recall = covered / float(len(sel))
        checks.append(_check("source_token_recall", "pass" if recall >= 0.35 else "fail",
                             "unit coverage >= 0.35", round(recall, 3),
                             "%d/%d plan-referenced units covered" % (covered, len(sel))))

    # --- numeric_preserved: direction is "deck numbers must exist in the source" ---
    # A slide deck is a *selective* product: failing because some source number never made
    # it onto a slide would fail every honest deck.  Coverage the other way round is
    # reported as informational (metrics.numeric_coverage) and never affects ``ok``.
    src_stats = set()
    for u in units:
        src_stats |= _stat_tokens(u.get("text") or "")
    for f in (inv.get("figures") or []) + (inv.get("tables") or []):
        if isinstance(f, dict):
            src_stats |= _stat_tokens(f.get("caption") or "")
    if start_text_chunks:
        src_stats |= _stat_tokens("\n".join(start_text_chunks))
    deck_stats = set()
    for txt in deck_by_n.values():
        deck_stats |= _stat_tokens(txt)
    if not deck_stats:
        checks.append(_check("numeric_preserved", "skipped", "statistics on the deck", 0,
                             "deck contains no statistic-bearing numeric token"))
    else:
        invented = sorted(deck_stats - src_stats)
        checks.append(_check("numeric_preserved", "fail" if invented else "pass",
                             "every statistic on the deck appears in the source material",
                             len(invented),
                             "deck statistic tokens=%d invented=%d (decimals/percent/units/"
                             "years/4+digit integers; structural integers excluded)"
                             % (len(deck_stats), len(invented))))
        for m in invented[:20]:
            violations.append({"kind": "numeric_invented", "slide": None, "text": m,
                               "why": "statistic on the deck absent from the source material"})
    not_on_deck = sorted(src_stats - deck_stats)
    metrics["numeric_coverage"] = (round((len(src_stats) - len(not_on_deck)) / float(len(src_stats)), 3)
                                   if src_stats else None)
    metrics["source_statistics_not_on_deck"] = not_on_deck
    if not_on_deck:
        informational.append({"kind": "numeric_not_on_deck",
                              "why": "source statistics not selected for any slide (informational)",
                              "count": len(not_on_deck), "sample": not_on_deck[:20]})
    deck_units, src_units = set(), set()
    for txt in deck_by_n.values():
        deck_units |= _unit_tokens(txt)
    for u in units:
        src_units |= _unit_tokens(u.get("text") or "")
    odd_units = sorted(deck_units - src_units)
    metrics["units_not_in_source"] = odd_units
    if odd_units:
        informational.append({"kind": "unit_not_in_source",
                              "why": "unit wording on the deck absent from the source (informational)",
                              "sample": odd_units[:20]})

    # --- figure_numbering: every 图N/表N used must exist in source labels / numbering map ---
    fig_labels = set()
    for key in ("figures", "tables"):
        for it in inv.get(key) or []:
            if isinstance(it, dict):
                fig_labels |= set(re.findall(r"(\d+)", str(it.get("label") or "")))
    numbering = plan.get("numbering") or {}
    for m in numbering.get("map") or []:
        if isinstance(m, str):
            mm = re.search(r"=.*?(\d+)", m)
            if mm:
                fig_labels.add(mm.group(1))
    refs = []
    for txt in list(deck_by_n.values()) + list(deck_by_n.values()):
        for mm in re.finditer(r"(?:图|表|Figure|Table|Fig\.?)\s*(\d+)", txt):
            refs.append((mm.group(1), mm.group(0)))
    all_refs = []
    for sl in deck:
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            for mm in re.finditer(r"(?:图|表|Figure|Table|Fig\.?)\s*(\d+)", sh["text"] or ""):
                all_refs.append((sl["n"], mm.group(1), mm.group(0)))
    metrics["figure_refs_checked"] = len(all_refs)
    bad_refs = [r for r in all_refs if r[1] not in fig_labels]
    metrics["figure_refs_ok"] = len(all_refs) - len(bad_refs)
    if not all_refs:
        checks.append(_check("figure_numbering", "skipped", "figure refs", 0,
                             "deck contains no 图N/表N references"))
    else:
        checks.append(_check("figure_numbering", "fail" if bad_refs else "pass",
                             "all refs exist in source labels", len(bad_refs),
                             "refs=%d bad=%d" % (len(all_refs), len(bad_refs))))
        for sl_n, num, tok in bad_refs:
            violations.append({"kind": "figure_number_unknown", "slide": sl_n,
                               "text": tok, "why": "no such label in source_inventory"})

    # --- caption_text_exact: captions must match source captions token-wise ---
    src_caps = []
    for it in (inv.get("figures") or []) + (inv.get("tables") or []):
        if isinstance(it, dict) and it.get("caption"):
            src_caps.append(it)
    cap_checked = 0
    cap_bad = []
    for sl in deck:
        for sh in sl["shapes"]:
            # structural caption detection only (p:ph or the agreed shape-name convention),
            # never "the text happens to start with a figure number"
            if sh["kind"] != "sp" or sh.get("role") != "caption":
                continue
            if not re.search(r"(?:图|表|Figure|Table)\s*\d+", sh["text"] or ""):
                continue
            toks = set(_content_tokens(sh["text"]))
            if not toks:
                continue
            best = 0.0
            for it in src_caps:
                ct = set(_content_tokens(it.get("caption") or ""))
                if ct:
                    best = max(best, len(toks & ct) / float(len(toks)))
            cap_checked += 1
            if best < 0.5:
                cap_bad.append({"slide": sl["n"], "text": (sh["text"] or "")[:120],
                                "overlap": round(best, 3)})
    if not src_caps:
        checks.append(_check("caption_text_exact", "skipped", "source captions", 0,
                             "source_inventory has no captions"))
    elif not cap_checked:
        checks.append(_check("caption_text_exact", "fail", ">=1 caption on deck", 0,
                             "no shape resolved to role=caption "
                             "(checked <p:ph> then the shape-name convention)"))
    else:
        checks.append(_check("caption_text_exact", "fail" if cap_bad else "pass",
                             "caption token overlap >= 0.5", len(cap_bad),
                             "captions checked=%d mismatched=%d" % (cap_checked, len(cap_bad))))
        for c in cap_bad:
            violations.append({"kind": "caption_mismatch", "slide": c["slide"],
                               "text": c["text"], "why": "overlap %.2f" % c["overlap"]})

    # --- notes_sourced ---
    notes_paras = []
    for sl in deck:
        nt = (sl.get("notes_text") or "").strip()
        if nt:
            notes_paras.append((sl["n"], nt))
    nbad = []
    for n, nt in notes_paras:
        toks = _content_tokens(nt)
        if len(toks) < 4:
            continue
        bad = [t for t in toks if t not in allowed]
        if bad and len(bad) / float(len(toks)) >= 0.6:
            nbad.append(n)
    if not notes_paras:
        checks.append(_check("notes_sourced", "skipped", "notes present", 0,
                             "deck has no notes text"))
    else:
        checks.append(_check("notes_sourced", "fail" if nbad else "pass",
                             "notes traceable to source", len(nbad),
                             "notes paragraphs=%d unsourced=%d" % (len(notes_paras), len(nbad))))
        for n in nbad:
            violations.append({"kind": "unsourced_notes", "slide": n, "text": "",
                               "why": "notes text absent from source"})

    # --- v1.3: start_deck_preserved / table_rendered / notes_evidence_present ---
    for helper, args in (("start_deck", (deck, plan, work)),
                         ("table", (deck, plan)),
                         ("notes", (deck, plan, inv, start_spec))):
        if helper == "start_deck":
            chk, vio, info, na = _fidelity_start_deck(*args)
        elif helper == "table":
            chk, vio, info, na = _fidelity_table(*args)
        else:
            chk, vio, info, na = _fidelity_notes(*args)
        if chk is not None:
            checks.append(chk)
        if na is not None:
            not_applicable.append(na)
            checks.append(_check(na["name"], "not_applicable", None, None, na["reason"]))
        violations.extend(vio)
        informational.extend(info)

    # --- no raw markup markers leaking into slide text ---
    markup_hits = []
    for sl in deck:
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            for para in sh["paras"]:
                t = para["text"] or ""
                for tag, rx in _MARKUP_RX:
                    if rx.search(t):
                        markup_hits.append({"slide": sl["n"], "marker": tag, "shape": sh.get("name"),
                                            "text": t[:100]})
                        break
    if not deck:
        checks.append(_check("no_markup_in_deck_text", "skipped", "slides present", 0,
                             "deck has no slides"))
    else:
        checks.append(_check("no_markup_in_deck_text", "fail" if markup_hits else "pass",
                             "no literal markdown/html markers in slide text",
                             len(markup_hits),
                             "markers checked=%d leaked=%d" % (len(_MARKUP_RX), len(markup_hits))))
        for h in markup_hits[:10]:
            violations.append({"kind": "markup_leak", "slide": h["slide"],
                               "text": h["text"], "why": "literal %s in slide text" % h["marker"]})

    vacuous = sum(1 for c in checks if c["status"] == "skipped")
    ok = all(c["status"] != "fail" for c in checks)
    if not ok and not violations:
        violations.append({"kind": "check_failed", "slide": None, "text": "",
                           "why": "a check failed without a violation entry"})
    demoted = 0
    if ok and violations:
        # invariant: fidelity.ok == true implies violations == []
        demoted = len(violations)
        informational.extend(violations)
        violations = []
    report = {
        "ok": ok,
        "checks": checks,
        "metrics": metrics,
        "violations": violations,
        "informational": informational,
        "invariant": {"ok_implies_no_violations": not (ok and violations),
                      "violations_demoted_to_informational": demoted},
        "unchecked": [{"name": c["name"], "reason": c["detail"]}
                      for c in checks if c["status"] == "skipped"],
        "not_applicable": not_applicable,
        "vacuous_skips": vacuous,
        "generated_at": _now(),
        "deck": None,
        "note": "`ok=false` is a hard failure: the delivery must not claim ok.",
    }
    return report


# ======================================================================================
# format audit
# ======================================================================================

CRITICAL = ("master_inheritance", "layout_link_resolves", "placeholder_idx_exists",
            "theme_fonts_preserved", "theme_colors_preserved", "slide_size_preserved",
            "text_no_overflow", "font_size_floor", "figure_aspect_ratio",
            "caption_number_consistency", "batch_consistency",
            "relationships_resolve", "content_types_default_namespace", "shape_overlap",
            "constraints_honoured")


def _resolve_template(work, manifest, spec):
    """Locate the template pptx bytes.

    Candidates, in order: the manifest input whose sha256 matches ``template_spec`` (its
    absolute path, then the byte-identical copy discover stages under ``<work>/inputs/``),
    the detected/template rel paths under ``<work>/inputs/``, and finally ``<work>`` itself.
    The candidate whose sha256 equals ``template_spec.sha256`` wins.
    """
    want = spec.get("sha256")
    cands = []
    for entry in (manifest.get("inputs") or []):
        if not isinstance(entry, dict) or entry.get("kind") not in ("pptx", "potx"):
            continue
        if entry.get("abs"):
            cands.append(entry["abs"])
        if entry.get("rel"):
            cands.append(os.path.join(work, "inputs", entry["rel"]))
    for rel in ((manifest.get("detected") or {}).get("template") or []):
        if isinstance(rel, str):
            cands.append(os.path.join(work, "inputs", rel))
    for rel in (spec.get("template_path"), spec.get("template")):
        if isinstance(rel, str) and rel:
            cands.append(os.path.join(work, "inputs", rel))
            cands.append(os.path.join(work, rel))
    cands.append(os.path.join(work, "template.pptx"))

    existing = [c for c in cands if c and os.path.isfile(c)]
    if want:
        for c in existing:
            try:
                if common.sha256_file(c) == want:
                    return c
            except Exception:
                continue
    # no sha match: still better to compare against a real template than to skip
    return existing[0] if existing else None


def run_format(work, output, parts, deck, manifest, spec, plan):
    checks = []
    inheritance = []
    overflow = []
    aspect = []
    not_applicable = []
    vacuous = 0

    def skip(name, reason):
        checks.append(_check(name, "skipped", None, None, reason))
        return 0

    # ---- slide size ----
    pres = None
    raw = parts.get("ppt/presentation.xml")
    if raw:
        try:
            pres = ET.fromstring(raw)
        except ET.ParseError:
            pres = None
    deck_size = None
    if pres is not None:
        sldsz = pres.find(_qn(P, "sldSz"))
        if sldsz is not None:
            try:
                deck_size = (int(sldsz.get("cx")), int(sldsz.get("cy")))
            except (TypeError, ValueError):
                deck_size = None
    want = spec.get("slide_size") or {}
    want_size = (want.get("cx"), want.get("cy"))
    if deck_size is None or want_size[0] is None:
        vacuous += 1
        skip("slide_size_preserved", "deck or template_spec has no slide size")
    else:
        ok = list(deck_size) == [want_size[0], want_size[1]]
        checks.append(_check("slide_size_preserved", "pass" if ok else "fail",
                             list(want_size), list(deck_size),
                             "ppt/presentation.xml/p:sldSz"))

    # ---- master inheritance: byte-for-byte, never "the part name exists" ----
    # "the template was really applied" is the headline claim of this task, so existence of
    # ppt/slideMasters/*.xml is not evidence: the bytes must match the template's.
    template_path = _resolve_template(work, manifest, spec)
    tparts_real = {}
    template_sha = None
    if template_path and os.path.isfile(template_path):
        template_sha = common.sha256_file(template_path)
        try:
            with zipfile.ZipFile(template_path) as z:
                tparts_real = {n: z.read(n) for n in z.namelist()}
        except Exception:
            tparts_real = {}

    def _is_inherited_part(name):
        return bool(re.match(r"^ppt/(slideMasters|slideLayouts|theme|notesMasters)/", name))

    tpl_inherited = sorted(n for n in tparts_real if _is_inherited_part(n))
    spec_needed = []
    for m in spec.get("masters") or []:
        if isinstance(m, dict):
            if m.get("part"):
                spec_needed.append(m["part"])
            if m.get("theme_rel"):
                spec_needed.append(m["theme_rel"])
            for lay in m.get("layouts") or []:
                spec_needed.append(lay)
    for t in spec.get("theme") or []:
        if isinstance(t, dict) and t.get("part"):
            spec_needed.append(t["part"])
    spec_needed = sorted(set(spec_needed))

    inherit_bytes = {"template": template_path, "template_sha256": template_sha,
                     "compared": 0, "identical": 0, "mismatched": [], "missing": []}
    if not tpl_inherited:
        vacuous += 1
        skip("master_inheritance",
             "template bytes unavailable for comparison (%s)"
             % ("template not found" if not template_path else "template unreadable/not a pptx"))
    else:
        for n in tpl_inherited:
            if n not in parts:
                inherit_bytes["missing"].append(n)
                continue
            inherit_bytes["compared"] += 1
            ts = common.sha256_bytes(tparts_real[n])
            ds = common.sha256_bytes(parts[n])
            if ts == ds:
                inherit_bytes["identical"] += 1
            else:
                inherit_bytes["mismatched"].append({"part": n, "template_sha256": ts,
                                                    "deck_sha256": ds})
        missing = [n for n in spec_needed if n not in parts]
        bad = inherit_bytes["missing"] or inherit_bytes["mismatched"] or missing
        checks.append(_check("master_inheritance", "fail" if bad else "pass",
                             "every inherited part (%d) is byte-identical to the template"
                             % len(tpl_inherited),
                             {"compared": inherit_bytes["compared"],
                              "identical": inherit_bytes["identical"],
                              "mismatched": len(inherit_bytes["mismatched"]),
                              "missing": len(inherit_bytes["missing"]) + len(missing)},
                             "template=%s sha256=%s; compared=%d identical=%d mismatched=%s "
                             "missing=%s"
                             % (template_path, (template_sha or "")[:16],
                                inherit_bytes["compared"], inherit_bytes["identical"],
                                [m["part"] for m in inherit_bytes["mismatched"][:5]],
                                (inherit_bytes["missing"] + missing)[:5]),
                             evidence=tpl_inherited[:12]))
        if inherit_bytes["mismatched"]:
            checks[-1]["detail_extra"] = inherit_bytes["mismatched"][:10]

    def _part_identical(name):
        if name and name in tparts_real and name in parts:
            return common.sha256_bytes(tparts_real[name]) == common.sha256_bytes(parts[name])
        return None

    for sl in deck:
        lp = sl.get("layout")
        mpart = None
        if lp and lp in parts:
            mrels = [x for x in _parse_rels(parts, lp) if x[1].endswith("/slideMaster")]
            mpart = mrels[0][4] if mrels and mrels[0][4] else None
        thpart = None
        if mpart and mpart in parts:
            trels = [x for x in _parse_rels(parts, mpart) if x[1].endswith("/theme")]
            thpart = trels[0][4] if trels and trels[0][4] else None
        per_part = {"layout": _part_identical(lp), "master": _part_identical(mpart),
                    "theme": _part_identical(thpart)}
        known = [v for v in per_part.values() if v is not None]
        inheritance.append({"slide": sl["n"], "layout": lp, "master": mpart, "theme": thpart,
                            "byte_identical": (all(known) if known else None),
                            "byte_identical_parts": per_part,
                            "placeholder_idx": [sh["ph"]["idx"] for sh in sl["shapes"]
                                                if sh.get("ph")]})

    # ---- layout link resolves (+ master/theme resolution) ----
    bad_layout = []
    layout_ph = {}
    for sl in deck:
        lp = sl["layout"]
        if not lp:
            bad_layout.append({"slide": sl["n"], "layout": None, "why": "slide has no layout rel"})
            continue
        if lp not in parts:
            bad_layout.append({"slide": sl["n"], "layout": lp, "why": "layout part missing"})
            continue
        layout_ph[sl["n"]] = _layout_ph_map(parts, lp)
        mrel = [x for x in _parse_rels(parts, lp) if x[1].endswith("/slideMaster")]
        if not mrel or not mrel[0][4] or mrel[0][4] not in parts:
            bad_layout.append({"slide": sl["n"], "layout": lp, "why": "layout master missing"})
        elif mrel[0][4]:
            trel = [x for x in _parse_rels(parts, mrel[0][4]) if x[1].endswith("/theme")]
            if not trel or not trel[0][4] or trel[0][4] not in parts:
                bad_layout.append({"slide": sl["n"], "layout": lp, "why": "master theme missing"})
    if not deck:
        vacuous += 1
        skip("layout_link_resolves", "deck has no slides")
    else:
        checks.append(_check("layout_link_resolves", "fail" if bad_layout else "pass",
                             "every slide -> existing layout -> master -> theme",
                             len(bad_layout), "unresolved=%d" % len(bad_layout)))
        for b in bad_layout:
            checks[-1].setdefault("detail_extra", [])
            checks[-1]["detail_extra"] = checks[-1].get("detail_extra", []) + [b]

    # ---- every internal relationship target must exist in the package ----
    rel_parts = sorted(n for n in parts if n.endswith(".rels"))
    dangling = []
    for rp in rel_parts:
        base_dir = os.path.dirname(rp)
        owner_dir = os.path.dirname(base_dir)
        owner_name = os.path.basename(rp)[:-5]
        owner_part = os.path.join(owner_dir, owner_name) if owner_name else ""
        try:
            root = ET.fromstring(parts[rp])
        except ET.ParseError:
            dangling.append({"rels": rp, "target": None, "why": "unparsable rels part"})
            continue
        for rel in root:
            if (rel.get("TargetMode") or "Internal") != "Internal":
                continue
            target = rel.get("Target") or ""
            resolved = _rel_target(owner_part, target)
            if resolved and resolved not in parts:
                dangling.append({"rels": rp, "target": target, "resolved": resolved})
    if not rel_parts:
        vacuous += 1
        skip("relationships_resolve", "package has no .rels parts")
    else:
        checks.append(_check("relationships_resolve", "fail" if dangling else "pass",
                             "every internal relationship target exists in the package",
                             len(dangling),
                             "rels parts=%d dangling=%d" % (len(rel_parts), len(dangling))))
        if dangling:
            checks[-1]["detail_extra"] = dangling[:10]

    # ---- [Content_Types].xml root must be an unprefixed <Types> element ----
    ct_raw = parts.get("[Content_Types].xml")
    if not ct_raw:
        vacuous += 1
        skip("content_types_default_namespace", "[Content_Types].xml absent from package")
    else:
        ct_text = ct_raw.decode("utf-8", "replace")
        body = re.sub(r"^\s*<\?xml[^>]*\?>", "", ct_text).lstrip()
        m = re.match(r"<([^\s/>]+)", body)
        root_name = m.group(1) if m else ""
        ns_ok = "http://schemas.openxmlformats.org/package/2006/content-types" in ct_text
        ct_ok = (root_name == "Types") and ns_ok
        checks.append(_check("content_types_default_namespace",
                             "pass" if ct_ok else "fail",
                             "root element is an unprefixed <Types> in the OPC namespace",
                             {"root": root_name, "ns_declared": ns_ok},
                             "a prefixed root parses fine in ElementTree/python-pptx but makes "
                             "LibreOffice reject the whole package"))

    # ---- placeholder idx exists ----
    bad_ph = []
    for sl in deck:
        mapping = layout_ph.get(sl["n"])
        if mapping is None:
            continue
        for sh in sl["shapes"]:
            if not sh.get("ph"):
                continue
            idx = sh["ph"]["idx"]
            if idx not in mapping:
                bad_ph.append({"slide": sl["n"], "idx": idx, "layout": sl["layout"]})
    if not deck:
        vacuous += 1
        skip("placeholder_idx_exists", "deck has no slides")
    else:
        checks.append(_check("placeholder_idx_exists", "fail" if bad_ph else "pass",
                             "every placed ph idx exists in its layout", len(bad_ph),
                             "unknown idx=%d" % len(bad_ph)))
        for b in bad_ph:
            checks[-1]["detail_extra"] = checks[-1].get("detail_extra", []) + [b]

    # ---- theme fonts / colors ----
    deck_axes = {(m.get("part") if isinstance(m, dict) else None) for m in (spec.get("masters") or [])}
    theme_parts = [t.get("part") for t in (spec.get("theme") or []) if isinstance(t, dict) and t.get("part")]
    if not theme_parts:
        theme_parts = _deck_theme_parts(parts)
    if not theme_parts:
        vacuous += 1
        skip("theme_fonts_preserved", "no theme part available")
        vacuous += 1
        skip("theme_colors_preserved", "no theme part available")
    else:
        font_bad, color_bad = [], []
        for t in (spec.get("theme") or []):
            if not isinstance(t, dict):
                continue
            got = _theme_of(parts, t.get("part"))
            if got is None:
                font_bad.append({"part": t.get("part"), "why": "theme part unreadable in deck"})
                color_bad.append({"part": t.get("part"), "why": "theme part unreadable in deck"})
                continue
            want_f = t.get("fonts") or {}
            for k, v in want_f.items():
                if v is not None and got["fonts"].get(k) != v:
                    font_bad.append({"part": t.get("part"), "key": k, "want": v,
                                     "got": got["fonts"].get(k)})
            want_c = t.get("colors") or {}
            for k, v in want_c.items():
                if v is None:
                    continue
                got_v = got["colors"].get(k)
                alts = got.get("colors_alt", {}).get(k) or []
                if got_v != v and v not in alts:
                    color_bad.append({"part": t.get("part"), "key": k, "want": v,
                                      "got": got_v, "alt": alts})
        checks.append(_check("theme_fonts_preserved", "fail" if font_bad else "pass",
                             "deck theme fonts == template theme fonts", len(font_bad),
                             "font mismatches=%d" % len(font_bad)))
        if font_bad:
            checks[-1]["detail_extra"] = font_bad[:10]
        checks.append(_check("theme_colors_preserved", "fail" if color_bad else "pass",
                             "deck theme colors == template theme colors", len(color_bad),
                             "color mismatches=%d" % len(color_bad)))
        if color_bad:
            checks[-1]["detail_extra"] = color_bad[:10]

    # ---- figure aspect ratio ----
    if not _midia(parts):
        vacuous += 1
        skip("figure_aspect_ratio", "deck package has no pictures")
    else:
        bad_ar = []
        pic_count = 0
        for sl in deck:
            for pic in sl["pics"]:
                x = pic.get("xfrm")
                tgt = pic.get("target")
                if not x or not tgt or tgt not in parts:
                    continue
                nat = common.natural_size(parts[tgt])
                if not nat or not nat[0] or not nat[1] or not x["cy"]:
                    continue
                pic_count += 1
                shape_ar = x["cx"] / float(x["cy"])
                img_ar = nat[0] / float(nat[1])
                delta = abs(shape_ar / img_ar - 1.0)
                scale = x["cx"] / float(nat[0] * 9525)  # 96 dpi reference
                entry = {"slide": sl["n"], "image": tgt,
                         "container": "%.2fx%.2fin" % (common.emu_to_in(x["cx"]),
                                                       common.emu_to_in(x["cy"])),
                         "scale": round(scale, 4), "aspect_delta": round(delta, 5)}
                aspect.append(entry)
                if delta > ASPECT_TOL:
                    bad_ar.append(entry)
        if not pic_count:
            vacuous += 1
            skip("figure_aspect_ratio", "no picture with a resolvable natural size")
        else:
            checks.append(_check("figure_aspect_ratio", "fail" if bad_ar else "pass",
                                 "aspect_delta <= %.2f" % ASPECT_TOL, len(bad_ar),
                                 "pictures measured=%d stretched=%d" % (pic_count, len(bad_ar))))

    # ---- caption / reference number consistency (same slide) ----
    bad_cap = []
    cap_checked = 0
    for sl in deck:
        body_nums, cap_nums = set(), set()
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            role = sh.get("role")
            if role in ("title", "meta", "pic"):
                continue
            txt = sh["text"] or ""
            nums = set(re.findall(r"(?:图|表|Figure|Table|Fig\.?)\s*(\d+)", txt))
            if not nums:
                continue
            if role == "caption":
                cap_nums |= nums
            else:
                body_nums |= nums
        if cap_nums:
            cap_checked += 1
            if body_nums and not (cap_nums & body_nums):
                bad_cap.append({"slide": sl["n"], "caption": sorted(cap_nums),
                                "body": sorted(body_nums)})
    if not cap_checked:
        vacuous += 1
        skip("caption_number_consistency",
             "no shape resolved to role=caption (checked <p:ph> then the shape-name "
             "convention: 'Figure Caption' -> caption)")
    else:
        checks.append(_check("caption_number_consistency", "fail" if bad_cap else "pass",
                             "caption number matches same-slide reference", len(bad_cap),
                             "captions checked=%d mismatched=%d" % (cap_checked, len(bad_cap))))

    # ---- text overflow (recomputed, never trusting adaptation_actions) ----
    floor_cfg = ((plan.get("font_ladder") or {}).get("floor") or {})
    layout_ph_all = {}
    for sl in deck:
        if sl.get("layout"):
            layout_ph_all[sl["n"]] = layout_ph.get(sl["n"]) or {}
    default_sz = {"title": 2800, "subtitle": 1800, "body": 1600, "caption": 1200, "table": 1200,
                  "meta": 1000}
    meas = 0
    for sl in deck:
        for sh in sl["shapes"]:
            if sh["kind"] != "sp" or sh.get("role") not in ("title", "subtitle", "body", "caption"):
                continue
            txt = sh["text"] or ""
            if not txt.strip():
                continue
            x = sh.get("xfrm")
            lp = layout_ph_all.get(sl["n"], {})
            ph = sh.get("ph") or {}
            if x is None and ph.get("idx") in lp:
                x = lp[ph["idx"]].get("xfrm")
            if not x or not x["cx"] or not x["cy"]:
                continue
            sizes = [s for p in sh["paras"] for s in p["sizes"]]
            if sizes:
                sz = max(sizes)
            elif ph.get("idx") in lp and lp[ph["idx"]].get("sz"):
                sz = lp[ph["idx"]]["sz"]
            else:
                sz = default_sz.get(sh.get("role"), 1600)
            font_pt = sz / 100.0
            scale = sh.get("autofit")
            if not scale:
                scale = _layout_font_scale(layout_ph, sl["n"], ph)
            if scale:
                font_pt = font_pt * float(scale)
            lins = (sh["paras"][0]["lnSpc"] if sh["paras"] and sh["paras"][0]["lnSpc"] else 1.2)
            l, r, t, b = _insets(_bodypr(sh["txbody"]) if sh["txbody"] is not None else None)
            box_w = common.emu_to_pt(max(x["cx"] - l - r, 0))
            box_h = common.emu_to_pt(max(x["cy"] - t - b, 0))
            est = common.estimate_wrapped_lines(txt, font_pt, box_w, line_spacing=lins)
            cap = common.capacity_lines(box_h, font_pt, lins)
            meas += 1
            over = (est.get("fits") is False) or (cap <= 0 and est.get("lines", 0) >= 1) \
                or (est.get("lines", 0) > cap)
            if over:
                action = "font_ladder" if sh.get("autofit") else "unresolved"
                overflow.append({"slide": sl["n"], "shape": sh.get("name") or ph.get("type"),
                                 "lines": est.get("lines"), "capacity_lines": cap,
                                 "font_pt": round(font_pt, 2), "action": action})
    if not meas:
        vacuous += 1
        skip("text_no_overflow", "no measurable text shape with an explicit/inherited box")
    else:
        unresolved = [o for o in overflow if o["action"] == "unresolved"]
        checks.append(_check("text_no_overflow", "fail" if unresolved else "pass",
                             "no unresolved overflow", len(unresolved),
                             "shapes measured=%d overflow=%d" % (meas, len(overflow))))

    # ---- font size floor ----
    # The *effective* size is what the audience sees: a run declared 18pt inside a shape
    # whose <a:bodyPr><a:normAutofit fontScale="50000"/> shrinks it renders at 9pt.
    # Squeezing text unreadable to defeat overflow is exactly the failure mode this task
    # names, so the floor is applied to declared_sz * fontScale, never to declared_sz alone.
    # plan font_ladder floors are OOXML 1/100 pt; tolerate a pt value just in case
    def _as_100pt(v, default):
        try:
            v = int(v)
        except (TypeError, ValueError):
            return default
        return v * 100 if 0 < v < 100 else v

    floors = {"title": 2400, "subtitle": common.FONT_FLOOR["body"],
              "body": _as_100pt(floor_cfg.get("body"), common.FONT_FLOOR["body"]),
              "caption": _as_100pt(floor_cfg.get("caption"), common.FONT_FLOOR["caption"]),
              "table": common.FONT_FLOOR["table"], "meta": 800}
    viol = []
    runs_checked = 0
    scaled_shapes = []
    eff_runs = []

    for sl in deck:
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            role = sh.get("role")
            if role not in floors:
                continue
            scale = sh.get("autofit")
            if not scale:
                scale = _layout_font_scale(layout_ph, sl["n"], sh.get("ph"))
            scale = float(scale or 1.0)
            if scale != 1.0:
                scaled_shapes.append({"slide": sl["n"], "shape": sh.get("name"),
                                      "font_scale": scale})
            for para in sh["paras"]:
                for sz in para["sizes"]:
                    runs_checked += 1
                    eff = sz * scale
                    eff_runs.append({"slide": sl["n"], "role": role, "eff_1_100pt": eff})
                    if eff < floors[role]:
                        viol.append({"slide": sl["n"], "shape": sh.get("name"), "role": role,
                                     "declared_sz": sz, "font_scale": round(scale, 5),
                                     "effective_pt": round(eff / 100.0, 2),
                                     "floor_pt": round(floors[role] / 100.0, 2),
                                     "declared_sz_1_100pt": sz,
                                     "floor_1_100pt": floors[role]})
    if not runs_checked:
        vacuous += 1
        skip("font_size_floor", "no explicit run sizes on the deck")
    else:
        checks.append(_check("font_size_floor", "fail" if viol else "pass",
                             "effective size (declared_sz x normAutofit fontScale) >= role floor",
                             len(viol),
                             "runs checked=%d below floor=%d; shapes with fontScale=%d; "
                             "effective size = sz x fontScale/100000"
                             % (runs_checked, len(viol), len(scaled_shapes))))
        if viol:
            checks[-1]["detail_extra"] = sorted(viol, key=lambda v: v["effective_pt"])[:10]
        if scaled_shapes:
            checks[-1]["font_scale_applied"] = scaled_shapes[:10]

    # ---- constraints_honoured (v1.3 ADDENDUM §A5) ----
    # "We read case.json.constraints" must be provable, not asserted: the run has to carry
    # the verbatim values into the plan and the produced deck has to obey the hard cap and
    # the readable-size floor.
    cons, cons_case_path = _case_constraints(work, manifest)
    applied = plan.get("constraints_applied") if isinstance(plan.get("constraints_applied"), dict) else {}
    if not _plan_is_v13(plan):
        reason = ("plan has no v1.3 `mode` key (legacy schema): the run cannot prove it read "
                  "case.json.constraints (no verbatim copy / constraints_applied)")
        not_applicable.append({"name": "constraints_honoured", "reason": reason})
        checks.append(_check("constraints_honoured", "not_applicable", None, None, reason))
    elif not cons and not applied:
        reason = "case.json provides no constraints and the plan records none"
        not_applicable.append({"name": "constraints_honoured", "reason": reason})
        checks.append(_check("constraints_honoured", "not_applicable", None, None, reason))
    else:
        subs, problems = [], []

        def _sub(name, ok, expected, observed, why=None):
            subs.append({"name": name, "ok": bool(ok), "expected": expected, "observed": observed})
            if not ok:
                problems.append(why or "%s: expected %r observed %r"
                                % (name, expected, observed))

        tgt = cons.get("target_slide_count")
        if isinstance(tgt, (int, float)) and not isinstance(tgt, bool) and int(tgt) > 0:
            tgt = int(tgt)
            _sub("slide_count_cap", len(deck) <= tgt, tgt, len(deck),
                 "slide_count %d exceeds target_slide_count %d" % (len(deck), tgt))
            pv = plan.get("target_slide_count")
            _sub("target_slide_count_verbatim", pv == tgt, tgt, pv,
                 "plan.target_slide_count=%r is not case.json.constraints.target_slide_count=%r"
                 % (pv, tgt))
        aud = cons.get("audience")
        if isinstance(aud, str) and aud.strip():
            pv = plan.get("audience")
            _sub("audience", isinstance(pv, str) and pv.strip() == aud.strip(), aud[:80], pv,
                 "plan.audience=%r differs from case.json.constraints.audience=%r" % (pv, aud))
        tm = cons.get("time_minutes")
        if isinstance(tm, (int, float)) and not isinstance(tm, bool):
            pv = plan.get("time_minutes")
            try:
                ok = pv is not None and abs(float(pv) - float(tm)) < 1e-6
            except (TypeError, ValueError):
                ok = False
            _sub("time_minutes", ok, tm, pv,
                 "plan.time_minutes=%r differs from case.json.constraints.time_minutes=%r"
                 % (pv, tm))
        case_floor = cons.get("font_size_floor_pt")
        if isinstance(case_floor, dict):
            plan_floor = ((plan.get("font_ladder") or {}).get("floor") or {})
            applied_floor = None
            if isinstance(applied.get("font_size_floor_pt"), dict):
                value = applied["font_size_floor_pt"].get("value")
                applied_floor = value if isinstance(value, dict) else None
            for role, pt in sorted(case_floor.items()):
                want = _as_100pt(pt, common.FONT_FLOOR.get(role))
                got = _as_100pt(plan_floor.get(role), None)
                if got is None and isinstance(applied_floor, dict):
                    got = _as_100pt(applied_floor.get(role), None)
                _sub("font_floor_%s" % role, got is not None and got >= want,
                     ">= %d (1/100 pt)" % want, got,
                     "plan font floor for %r is %r, below case.json floor %r"
                     % (role, got, want))
                below = [r for r in eff_runs if r["role"] == role and r["eff_1_100pt"] < want]
                _sub("font_floor_runs_%s" % role, not below, 0, len(below),
                     "%d rendered %r run(s) below case.json floor %.2f pt"
                     % (len(below), role, want / 100.0))
        cn = cons.get("caption_numbering")
        if isinstance(cn, str) and cn.strip():
            pv = (plan.get("numbering") or {}).get("mode")
            _sub("caption_numbering", pv == cn, cn, pv,
                 "plan.numbering.mode=%r differs from case.json.constraints.caption_numbering=%r"
                 % (pv, cn))
        if not subs:
            reason = "case.json.constraints carries none of audience/time_minutes/" \
                     "target_slide_count/font_size_floor_pt/caption_numbering"
            not_applicable.append({"name": "constraints_honoured", "reason": reason})
            checks.append(_check("constraints_honoured", "not_applicable", None, None, reason))
        else:
            chk = _check("constraints_honoured", "fail" if problems else "pass",
                         "case.json.constraints honoured and carried verbatim into the plan",
                         len(problems),
                         "subchecks=%d violated=%d (case.json=%s)"
                         % (len(subs), len(problems), cons_case_path))
            chk["subchecks"] = subs
            if problems:
                chk["detail_extra"] = problems[:10]
            checks.append(chk)

    # ---- batch consistency ----
    groups = {}
    for sl in deck:
        groups.setdefault(sl.get("layout"), []).append(sl)
    inconsistent = []
    for lp, slides in groups.items():
        if len(slides) < 2:
            continue
        titled = []
        for sl in slides:
            for sh in sl["shapes"]:
                if sh.get("role") == "title" and sh.get("xfrm"):
                    titled.append((sl["n"], sh["xfrm"]))
                    break
        if len(titled) >= 2:
            ref = titled[0][1]
            for n, x in titled[1:]:
                if (abs(x["x"] - ref["x"]) > 1270 or abs(x["y"] - ref["y"]) > 1270):
                    inconsistent.append({"layout": lp, "slide": n,
                                         "why": "title x/y differs from sibling slide"})
        nums = {(sl["slide_num_fld"], sl["footer_fld"]) for sl in slides}
        if len(nums) > 1:
            inconsistent.append({"layout": lp, "slide": None,
                                 "why": "slide-number/footer presence differs within layout"})
    size_viol = []
    for sl in deck:
        tsz = max([s for sh in sl["shapes"] if sh.get("role") == "title"
                   for p in sh["paras"] for s in p["sizes"]] or [0])
        bsz = max([s for sh in sl["shapes"] if sh.get("role") == "body"
                   for p in sh["paras"] for s in p["sizes"]] or [0])
        if tsz and bsz and tsz < bsz:
            size_viol.append({"slide": sl["n"], "title": tsz, "body": bsz})
    if not deck:
        vacuous += 1
        skip("batch_consistency", "deck has no slides")
    else:
        fails = len(inconsistent) + len(size_viol)
        checks.append(_check("batch_consistency", "fail" if fails else "pass",
                             "title position / numbering / size hierarchy consistent",
                             fails, "layout inconsistencies=%d size-hierarchy=%d"
                             % (len(inconsistent), len(size_viol))))

    # ---- shape overlap: two text-bearing shapes printed on top of each other ----
    # Complements text_no_overflow: overflow is "text leaves its own box", overlap is
    # "two boxes intersect".  Rects inherited from the same layout placeholder are a
    # legitimate co-location and are skipped.
    overlap = []
    dup_notes = []
    pairs_checked = 0
    master_ph_cache = {}
    for sl in deck:
        pmap = layout_ph.get(sl["n"]) or {}
        mmap = {}
        if sl.get("layout"):
            mrels = [x for x in _parse_rels(parts, sl["layout"]) if x[1].endswith("/slideMaster")]
            mpart = mrels[0][4] if mrels and mrels[0][4] else None
            if mpart:
                if mpart not in master_ph_cache:
                    master_ph_cache[mpart] = _master_ph_map(parts, mpart)
                mmap = master_ph_cache[mpart]
        items = []
        seen_ph = set()
        for sh in sl["shapes"]:
            if sh["kind"] != "sp":
                continue
            text = "\n".join(p["text"] for p in sh["paras"] if (p["text"] or "").strip())
            if not text.strip():
                continue
            ph = sh.get("ph") or {}
            rect = sh.get("xfrm")
            if rect is None and ph.get("idx") in pmap:
                rect = pmap[ph["idx"]].get("xfrm")
            if rect is None and ph.get("idx") in mmap:
                # placeholders commonly inherit their rectangle from the master
                rect = mmap[ph["idx"]].get("xfrm")
            if rect is None or not rect.get("cx") or not rect.get("cy"):
                continue
            key = (ph.get("type"), ph.get("idx")) if ph else None
            if key is not None:
                if key in seen_ph:
                    dup_notes.append({"slide": sl["n"], "placeholder": list(key)})
                seen_ph.add(key)
            items.append({"name": sh.get("name") or (ph.get("type") or "shape"),
                          "rect": rect, "ph_key": key})
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                a, b = items[i], items[j]
                if a["ph_key"] and a["ph_key"] == b["ph_key"]:
                    continue
                ra, rb = a["rect"], b["rect"]
                pairs_checked += 1
                ix = min(ra["x"] + ra["cx"], rb["x"] + rb["cx"]) - max(ra["x"], rb["x"])
                iy = min(ra["y"] + ra["cy"], rb["y"] + rb["cy"]) - max(ra["y"], rb["y"])
                if ix <= 0 or iy <= 0:
                    continue
                denom = min(ra["cx"] * ra["cy"], rb["cx"] * rb["cy"]) or 1
                ratio = (ix * iy) / float(denom)
                if ratio > OVERLAP_TOL:
                    overlap.append({"slide": sl["n"], "a": a["name"], "b": b["name"],
                                    "overlap_ratio": round(ratio, 3),
                                    "rects": {"a": ra, "b": rb}})
    if not pairs_checked:
        vacuous += 1
        skip("shape_overlap", "no two text-bearing shapes with resolvable rectangles")
    else:
        checks.append(_check("shape_overlap", "fail" if overlap else "pass",
                             "no two text shapes overlap > %.2f of the smaller area" % OVERLAP_TOL,
                             len(overlap),
                             "overlapping pairs=%d of %d pairs checked; duplicate placeholders=%d"
                             % (len(overlap), pairs_checked, len(dup_notes))))
        if overlap:
            checks[-1]["detail_extra"] = overlap[:10]

    role_sources = {}
    for sl in deck:
        for sh in sl["shapes"]:
            key = (sh.get("role_source") or "default").split(":")[0]
            role_sources[key] = role_sources.get(key, 0) + 1

    has_fail = any(c["status"] == "fail" for c in checks)
    has_skip = any(c["status"] == "skipped" for c in checks)
    status = "fail" if has_fail else ("partial" if has_skip else "pass")
    report = {
        "status": status,
        "checks": checks,
        "overflow": overflow,
        "aspect": aspect,
        "overlap": overlap,
        "duplicate_placeholders": dup_notes,
        "inheritance": inheritance,
        "inheritance_bytes": inherit_bytes,
        "unchecked": [{"name": c["name"], "reason": c["detail"]}
                      for c in checks if c["status"] == "skipped"],
        "not_applicable": not_applicable,
        "role_resolution": role_sources,
        "font_scale": {"shapes_with_font_scale": scaled_shapes[:40],
                       "count": len(scaled_shapes),
                       "note": "font_size_floor and text_no_overflow use declared_sz x "
                               "normAutofit fontScale/100000 as the effective size"},
        "vacuous_skips": vacuous,
        "generated_at": _now(),
        "note": "skipped != pass; unchecked items never count as PASS.",
    }
    return report


# ======================================================================================
# entry point
# ======================================================================================

def _find_deck(output_dir, plan):
    name = (plan or {}).get("deck_name") or common.DECK_DEFAULT_NAME
    cand = os.path.join(output_dir, name)
    if os.path.isfile(cand):
        return cand
    found = []
    for root, dirs, files in os.walk(output_dir):
        dirs[:] = [d for d in dirs if d not in ("preview", "__pycache__")]
        for f in sorted(files):
            if f.lower().endswith((".pptx", ".pptm")) and not f.startswith("~$"):
                found.append(os.path.join(root, f))
    if not found:
        return None
    for f in found:
        if os.path.basename(f) == name:
            return f
    return found[0]


def verify(work_dir: str, output_dir: str) -> dict:
    """Audit <output>/<deck> against <work> manifests. Writes both reports."""
    work_dir = os.path.abspath(work_dir)
    output_dir = os.path.abspath(output_dir)
    manifest = _read_json(os.path.join(work_dir, "case_manifest.json"))
    spec = _read_json(os.path.join(work_dir, "template_spec.json"))
    inv = _read_json(os.path.join(work_dir, "source_inventory.json"))
    plan = _read_json(os.path.join(work_dir, "slides_plan.json"))

    missing_inputs = [n for n in WORK_REQUIRED
                      if not os.path.isfile(os.path.join(work_dir, n))]
    deck_path = _find_deck(output_dir, plan)
    if deck_path is None:
        missing_inputs.append("<output>/*.pptx")

    if missing_inputs:
        reason = "missing required inputs: %s" % ", ".join(missing_inputs)
        check = _check("inputs_present", "fail", "all work/deck inputs", len(missing_inputs), reason)
        fid = {"ok": False, "checks": [check], "metrics": {}, "violations": [],
               "unchecked": [], "not_applicable": [], "vacuous_skips": 0, "generated_at": _now(),
               "deck": deck_path, "note": reason}
        fmt = {"status": "fail", "checks": [check], "overflow": [], "aspect": [],
               "inheritance": [], "unchecked": [], "not_applicable": [], "vacuous_skips": 0,
               "generated_at": _now(), "note": reason}
        common.write_json(os.path.join(work_dir, "fidelity_report.json"), fid)
        common.write_json(os.path.join(work_dir, "format_audit.json"), fmt)
        return {"ok": False, "reason": reason, "fidelity": fid, "format": fmt,
                "fidelity_report": os.path.join(work_dir, "fidelity_report.json"),
                "format_audit": os.path.join(work_dir, "format_audit.json"),
                "deck": deck_path, "vacuous_skips": 0}

    try:
        with zipfile.ZipFile(deck_path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
    except Exception as exc:  # noqa: BLE001
        reason = "deck is not a readable zip: %s: %s" % (type(exc).__name__, exc)
        check = _check("deck_readable", "fail", "readable pptx zip", 0, reason)
        fid = {"ok": False, "checks": [check], "metrics": {}, "violations": [],
               "unchecked": [], "not_applicable": [], "vacuous_skips": 0, "generated_at": _now(),
               "deck": deck_path, "note": reason}
        fmt = {"status": "fail", "checks": [check], "overflow": [], "aspect": [],
               "inheritance": [], "unchecked": [], "not_applicable": [], "vacuous_skips": 0,
               "generated_at": _now(), "note": reason}
        common.write_json(os.path.join(work_dir, "fidelity_report.json"), fid)
        common.write_json(os.path.join(work_dir, "format_audit.json"), fmt)
        return {"ok": False, "reason": reason, "fidelity": fid, "format": fmt,
                "fidelity_report": os.path.join(work_dir, "fidelity_report.json"),
                "format_audit": os.path.join(work_dir, "format_audit.json"),
                "deck": deck_path, "vacuous_skips": 0}

    deck = _extract_deck(parts, spec)
    fid = run_fidelity(work_dir, parts, deck, manifest or {}, inv or {}, plan or {})
    fmt = run_format(work_dir, output_dir, parts, deck, manifest or {}, spec or {}, plan or {})
    fid["deck"] = common.relpath_or_abspath(deck_path, work_dir)
    fid["deck_sha256"] = common.sha256_file(deck_path)
    fid["slide_count"] = len(deck)
    fmt["deck"] = fid["deck"]
    fmt["slide_count"] = len(deck)
    common.write_json(os.path.join(work_dir, "fidelity_report.json"), fid)
    common.write_json(os.path.join(work_dir, "format_audit.json"), fmt)

    v13_status = {}
    for name in V13_CHECKS:
        status = None
        for c in list(fid.get("checks") or []) + list(fmt.get("checks") or []):
            if c.get("name") == name:
                status = c.get("status")
                break
        v13_status[name] = status
    na = (fid.get("not_applicable") or []) + (fmt.get("not_applicable") or [])
    return {"ok": bool(fid.get("ok")) and fmt.get("status") != "fail",
            "fidelity_ok": bool(fid.get("ok")), "format_status": fmt.get("status"),
            "fidelity": fid, "format": fmt,
            "fidelity_report": os.path.join(work_dir, "fidelity_report.json"),
            "format_audit": os.path.join(work_dir, "format_audit.json"),
            "deck": fid["deck"], "slide_count": len(deck),
            "v13": v13_status, "not_applicable": na,
            "vacuous_skips": int(fid.get("vacuous_skips", 0)) + int(fmt.get("vacuous_skips", 0))}


def main(argv=None):
    ap = argparse.ArgumentParser(description="independent pptx fidelity/format audit")
    ap.add_argument("--work", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args(argv)
    result = verify(args.work, args.output)
    summary = {"ok": result.get("ok"), "fidelity_ok": result.get("fidelity_ok"),
               "format_status": result.get("format_status"),
               "slide_count": result.get("slide_count"),
               "vacuous_skips": result.get("vacuous_skips"),
               "reason": result.get("reason")}
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
