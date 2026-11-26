#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""discover.py — evidence-first reconnaissance for the academic-PPT template task.

Reads the real case input (``case.json`` + ``task.md`` + materials) and writes, into the
work directory:

* ``case_manifest.json``    — inputs, original ``output_contract``, required dotted leaves
* ``template_spec.json``    — sldSz / masters / layouts / placeholders / theme, byte-derived
* ``source_inventory.json`` — sections, units (with locators), figures, tables
* ``slide_catalog.jsonl``   — one record per source unit, reuseable by the planner
* ``REVIEW_PACKET.md``      — human entry point (<=200 lines)
* ``sources.json``          — provenance / toolchain / reproduce commands (initial version)

Everything here is *read* from real bytes: nothing is filled from domain defaults, and a
missing or unreadable input degrades to an honest note instead of an invented value.

Public entry point (called by ``run.py prepare``)::

    discover(input_dir, work_dir, output_dir) -> dict

Standalone::

    python -B discover.py --input <input> --work <work> [--output <output>]

Pure standard library, offline, POSIX/Windows safe, never writes ``__pycache__``.
Owner: engine (see work/CONTRACT.md §1/§5).
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import zipfile
from xml.etree import ElementTree as ET

sys.dont_write_bytecode = True
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import common  # noqa: E402  (frozen shared layer, CONTRACT §3)

NS = common.NS
P = NS["p"]
A = NS["a"]
R = NS["r"]

REL_OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

ROLE_ALIASES = {
    "template": ("template", "templ", "模板", "母版模板", "格式范本", "样式模板"),
    "source": ("source", "paper", "manuscript", "original", "材料", "论文", "正文稿", "源文件", "稿件"),
    "start_deck": ("start_deck", "startdeck", "start-deck", "starting_deck", "existing_deck",
                   "deck", "起始", "已有演示", "待改造"),
    "figures": ("figure", "fig", "image", "picture", "图", "插图", "图片"),
}

TEXT_KINDS = ("md", "txt", "pdf", "docx", "tex", "other")
PPTX_KINDS = ("pptx", "potx", "pptm")
IMAGE_KINDS = ("png", "jpg", "jpeg", "gif", "bmp", "tiff", "webp", "svg", "emf", "wmf")

DEFAULT_BUDGET = {"model_requests": 64, "seconds": 900}

# Caption recognition (source numbering is preserved; these only *read* the label).
# A caption is the *sentence* that carries the number, never the whole paragraph:
# result paragraphs routinely mention 图1, 图2 and 表1 in the same block.
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_LABEL_RE = re.compile(r"(图|表|Figure|Fig\.?|Table|Tab\.?)\s*([0-9]+|[一二三四五六七八九十]+)",
                       re.IGNORECASE)
_SENT_BREAK = "。！？!?；;"
_CAPTION_MAX_CHARS = 160


def _normalise_label(raw):
    """``Fig. 1`` -> ``Figure 1``; ``图 1`` -> ``图1`` (same label across references)."""
    text = re.sub(r"\s+", " ", (raw or "").strip())
    text = re.sub(r"^Fig(?:ure|\.)?\s*", "Figure ", text, flags=re.IGNORECASE)
    text = re.sub(r"^Tab(?:le|\.)?\s*", "Table ", text, flags=re.IGNORECASE)
    text = re.sub(r"^(图|表)\s*", r"\1", text)
    return text


def _label_occurrences(text):
    """``[(kind, label, char_offset)]`` for every figure/table number mentioned in ``text``."""
    out = []
    for match in _LABEL_RE.finditer(text or ""):
        raw = match.group(0)
        kind = "figure" if re.match(r"^(图|Figure|Fig)", raw, re.IGNORECASE) else "table"
        out.append((kind, _normalise_label(raw), match.start()))
    return out


def _label_pattern(label):
    body = re.escape(label).replace(r"\ ", r"\s*")
    # 图1 must not match inside 图12; Figure 1 must not match Figure 10
    return re.compile(r"(?<![0-9A-Za-z])" + body + r"(?![0-9])", re.IGNORECASE)


def _sentence_with_label(block, label):
    """First sentence (split on 。！？.!?；; and newlines) that carries ``label``."""
    pattern = _label_pattern(label)
    for part in re.split(r"(?<=[%s])" % re.escape(_SENT_BREAK), block or ""):
        if pattern.search(part):
            return part.strip()
    return None


def _trim_caption(sentence, label):
    """Reduce a sentence to the clause that belongs to ``label``.

    The caption never contains another figure/table number: text before the first
    occurrence of *another* number is dropped (a "见表1 与 Figure 3" sentence yields
    ``表1 与`` for 表1) and text after the next different number is dropped as well.
    """
    occurrences = _label_occurrences(sentence)
    own = [position for _kind, found, position in occurrences if found == label]
    if not own:
        return sentence.strip(), False
    earlier_other = [position for _kind, found, position in occurrences if found != label and position < own[0]]
    # normally keep the sentence from its own start; only when an earlier number owns the
    # prefix do we begin exactly at our own number
    start = own[0] if earlier_other else 0
    after = [position for _kind, found, position in occurrences if found != label and position > start]
    end = min(after) if after else len(sentence)
    sentence = sentence[start:end].strip(" \t\n，,、；;：:")
    truncated = False
    if len(sentence) > _CAPTION_MAX_CHARS:
        window = sentence[:_CAPTION_MAX_CHARS]
        cut = max((window.rfind(ch) for ch in _SENT_BREAK), default=-1)
        sentence = window[: cut + 1] if cut >= 0 else window
        truncated = True
    return sentence.strip(), truncated


def extract_caption_entries(block):
    """One entry per distinct figure/table number in ``block``, each with its own sentence.

    The paragraph itself is *not* consumed here: callers keep it as a ``text`` unit, so no
    body text is lost when a paragraph also introduces a figure.
    """
    entries = []
    seen = set()
    for kind, label, position in _label_occurrences(block):
        if (kind, label) in seen:
            continue
        seen.add((kind, label))
        sentence = _sentence_with_label(block, label)
        if sentence is None:
            # fall back to the clause that literally holds the number
            tail = block[position:]
            for index, char in enumerate(tail):
                if char in _SENT_BREAK or char == "\n":
                    tail = tail[: index + 1]
                    break
            sentence = tail.strip()
        caption, truncated = _trim_caption(sentence, label)
        if not caption:
            continue
        entries.append({
            "kind": kind,
            "label": label,
            "caption": caption,
            # the locator points at the character where the number itself appears
            "offset": position,
            "truncated": truncated,
        })
    return entries


def _record_caption(entry, unit_record, locator, source_ref, figures, tables):
    """Register (or merge a repeated reference into) one figure/table registry entry."""
    registry = figures if entry["kind"] == "figure" else tables
    existing = next((item for item in registry if item["label"] == entry["label"]), None)
    if existing is not None:
        # the same label referred to again elsewhere: keep one entry, record every real locator
        existing.setdefault("also_at", []).append({"source_ref": source_ref, "unit": unit_record["id"]})
        unit_record["registry_id"] = existing["id"]
        unit_record["duplicate_label_of"] = existing["id"]
        return existing["id"]
    item = {
        "id": ("F%d" if entry["kind"] == "figure" else "T%d") % (len(registry) + 1),
        "label": entry["label"],
        "caption": entry["caption"],
        "caption_plain": _strip_markup(entry["caption"]),
        "source_ref": source_ref,
        "locator": locator,
        "unit": unit_record["id"],
        "section": unit_record.get("section"),
    }
    if entry["kind"] == "figure":
        item.update({"asset": None, "width": None, "height": None})
    else:
        # v1.3: real cell data is attached later by ``_attach_table_data``; until then the
        # lists stay empty ("no invented cells") and ``data_note`` explains why.
        item.update({"header": [], "rows": [], "data_source": None, "data_note": None})
    if entry.get("truncated"):
        item["caption_truncated"] = True
    registry.append(item)
    return item["id"]


class DiscoveryError(Exception):
    """Input-level failure that still produces honest degraded artifacts."""


# --------------------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------------------

def _resolve_target(base_part, target):
    if not target:
        return None
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base_dir = posixpath.dirname(base_part) if base_part else ""
    return posixpath.normpath(posixpath.join(base_dir, target))


def _rels_part(part):
    if not part:
        return "_rels/.rels"
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _read_rels(zf, part):
    """Return ``{id: {'type','target','mode','part'}}`` for one part (resolved targets)."""
    name = _rels_part(part)
    if name not in zf.namelist():
        return {}
    try:
        root = ET.fromstring(zf.read(name))
    except ET.ParseError:
        return {}
    out = {}
    for el in root.iter():
        if common.localname(el.tag) != "Relationship":
            continue
        rid = el.get("Id")
        if not rid:
            continue
        target = el.get("Target")
        out[rid] = {
            "type": el.get("Type") or "",
            "target": target,
            "mode": el.get("TargetMode") or "Internal",
            "part": _resolve_target(part, target) if (el.get("TargetMode") or "Internal") != "External" else target,
        }
    return out


def _child(elem, tag):
    if elem is None:
        return None
    return elem.find(tag, NS)


def _attr(elem, name):
    if elem is None:
        return None
    return elem.get(name)


def _color_of(scheme_child):
    """Hex colour from ``<a:srgbClr val=..>`` or ``<a:sysClr lastClr=..>``."""
    if scheme_child is None:
        return None
    for el in scheme_child.iter():
        local = common.localname(el.tag)
        if local == "srgbClr" and el.get("val"):
            return el.get("val").upper()
        if local == "sysClr":
            val = el.get("lastClr") or el.get("val")
            if val:
                return val.upper()
    return None


def _xfrm_of(sp):
    """``(x, y, cx, cy)`` of a ``p:sp`` shape, or ``None`` when the shape has no xfrm."""
    for path in (("p:spPr", "a:xfrm"), ("p:xfrm",)):
        node = sp
        ok = True
        for step in path:
            node = _child(node, step)
            if node is None:
                ok = False
                break
        if not ok:
            continue
        off = _child(node, "a:off")
        ext = _child(node, "a:ext")
        if off is None and ext is None:
            continue
        try:
            x = int(off.get("x")) if off is not None and off.get("x") is not None else 0
            y = int(off.get("y")) if off is not None and off.get("y") is not None else 0
            cx = int(ext.get("cx")) if ext is not None and ext.get("cx") is not None else 0
            cy = int(ext.get("cy")) if ext is not None and ext.get("cy") is not None else 0
        except (TypeError, ValueError):
            continue
        return {"x": x, "y": y, "cx": cx, "cy": cy}
    return None


def _body_pr_of(sp):
    body = sp.find("p:txBody/a:bodyPr", NS)
    if body is None:
        return None, None
    sz = body.get("sz")
    try:
        sz = int(sz) if sz is not None else None
    except ValueError:
        sz = None
    autofit = body.find("a:normAutofit", NS)
    info = {
        "sz": sz,
        "anchor": body.get("anchor"),
        "normAutofit_fontScale": (autofit.get("fontScale") if autofit is not None else None),
        "attrs": {k: v for k, v in body.attrib.items()},
    }
    return info, body


def _placeholder_shapes(root):
    """Yield ``(sp, ph_element, cNvPr)`` for every placeholder shape inside a cSld spTree."""
    if root is None:
        return []
    out = []
    for sp in root.iter("{%s}sp" % P):
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None:
            continue
        cnv = sp.find("p:nvSpPr/p:cNvPr", NS)
        out.append((sp, ph, cnv))
    return out


def _ph_key(ph):
    """Effective ``(idx, type)`` — OOXML defaults: idx 0, type ``obj``."""
    idx = ph.get("idx")
    idx = str(idx) if idx is not None and str(idx) != "" else "0"
    ph_type = ph.get("type")
    ph_type = ph_type if ph_type else "obj"
    return idx, ph_type


def _ph_record(sp, ph, cnv):
    idx, ph_type = _ph_key(ph)
    xfrm = _xfrm_of(sp)
    body_pr, _ = _body_pr_of(sp)
    return {
        "idx": idx,
        "ph_type": ph_type,
        "name": (cnv.get("name") if cnv is not None else None),
        "xfrm": xfrm,
        "body_pr": body_pr,
        "raw": {"idx_attr": ph.get("idx"), "type_attr": ph.get("type")},
    }


# --------------------------------------------------------------------------------------
# pptx template spec (only from real bytes)
# --------------------------------------------------------------------------------------

def _parse_theme(zf, part):
    root = common.load_zip_xml(zf, part)
    if root is None:
        return None
    theme = {"part": part, "name": root.get("name"), "fonts": {}, "colors": {}}
    fonts = root.find("a:themeElements/a:fontScheme", NS)
    if fonts is not None:
        for which, key in (("majorFont", "major"), ("minorFont", "minor")):
            node = fonts.find("a:%s" % which, NS)
            if node is None:
                continue
            theme["fonts"]["%s_latin" % key] = _attr(node.find("a:latin", NS), "typeface")
            theme["fonts"]["%s_ea" % key] = _attr(node.find("a:ea", NS), "typeface")
            theme["fonts"]["%s_cs" % key] = _attr(node.find("a:cs", NS), "typeface")
    scheme = root.find("a:themeElements/a:clrScheme", NS)
    if scheme is not None:
        for child in list(scheme):
            local = common.localname(child.tag)
            theme["colors"][local] = _color_of(child)
    return theme


def _parse_master(zf, part):
    root = common.load_zip_xml(zf, part)
    if root is None:
        return None, []
    c_sld = _child(root, "p:cSld")
    master = {
        "part": part,
        "name": _attr(c_sld, "name"),
        "theme_rel": None,
        "layouts": [],
        "placeholders": [],
    }
    for sp, ph, cnv in _placeholder_shapes(c_sld):
        rec = _ph_record(sp, ph, cnv)
        master["placeholders"].append({
            "idx": rec["idx"], "ph_type": rec["ph_type"], "name": rec["name"], "xfrm": rec["xfrm"],
        })
    rels = _read_rels(zf, part)
    for rid, rel in rels.items():
        if rel["type"].endswith("/theme") and rel.get("part"):
            master["theme_rel"] = rel["part"]
    layout_ids = []
    id_lst = _child(root, "p:sldLayoutIdLst")
    if id_lst is not None:
        for el in id_lst:
            if common.localname(el.tag) != "sldLayoutId":
                continue
            rid = el.get("{%s}id" % R)
            rel = rels.get(rid or "")
            if rel and rel.get("part"):
                layout_ids.append((rid, rel["part"]))
            else:
                layout_ids.append((rid, None))
    return master, layout_ids


def _parse_layout(zf, part, master_part, master_index):
    root = common.load_zip_xml(zf, part)
    layout = {
        "part": part,
        "name": None,
        "type": None,
        "master": master_part,
        "placeholders": [],
        "note": None,
    }
    if root is None:
        layout["note"] = "layout part is not parseable XML"
        return layout
    c_sld = _child(root, "p:cSld")
    layout["name"] = _attr(c_sld, "name")
    layout["type"] = root.get("type") or "custom"
    for sp, ph, cnv in _placeholder_shapes(c_sld):
        rec = _ph_record(sp, ph, cnv)
        entry = {
            "idx": rec["idx"],
            "ph_type": rec["ph_type"],
            "name": rec["name"],
            "xfrm": rec["xfrm"],
            "body_pr": rec["body_pr"],
            "xfrm_source": "layout" if rec["xfrm"] else None,
            "inherits_from_master": rec["xfrm"] is None,
            "raw": rec["raw"],
        }
        if rec["xfrm"] is None:
            key = (rec["idx"], rec["ph_type"])
            master_ph = master_index.get(key)
            match_kind = "idx"
            if master_ph is None or not master_ph.get("xfrm"):
                # title/ctrTitle and the dt/ftr/sldNum placeholders inherit by placeholder
                # *type*: layouts use idx 10/11/12 while masters may use 2/3/4.
                for (midx, mtype), cand in master_index.items():
                    if mtype == rec["ph_type"] and cand.get("xfrm"):
                        master_ph = cand
                        match_kind = "type"
                        break
            if master_ph is not None and master_ph.get("xfrm"):
                entry["xfrm"] = master_ph.get("xfrm")
                entry["xfrm_source"] = "master"
                entry["xfrm_match"] = match_kind
            else:
                entry["xfrm_source"] = "unresolved"
                entry["xfrm_note"] = ("layout placeholder declares no xfrm and no master placeholder "
                                      "with a usable xfrm was found")
            entry["inherits_from_master"] = True
        layout["placeholders"].append(entry)
    return layout


def _content_type_overrides(zf):
    root = common.load_zip_xml(zf, "[Content_Types].xml")
    out = {}
    if root is None:
        return out
    for el in root.iter():
        if common.localname(el.tag) == "Override" and el.get("PartName"):
            out[el.get("PartName").lstrip("/")] = el.get("ContentType")
    return out


def parse_template(path, rel, sha):
    """Build ``template_spec`` strictly from the pptx bytes on disk."""
    spec = {
        "template_path": rel,
        "sha256": sha,
        "ok": False,
        "slide_size": None,
        "masters": [],
        "layouts": [],
        "theme": [],
        "notes_master": None,
        "evidence": [],
        "notes": [],
    }
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if "[Content_Types].xml" not in names or "_rels/.rels" not in names:
                raise DiscoveryError("package lacks [Content_Types].xml or _rels/.rels")
            bad = zf.testzip()
            if bad:
                raise DiscoveryError("CRC failure in member %s" % bad)
            root_rels = _read_rels(zf, "")
            pres_part = None
            for rel_entry in root_rels.values():
                if rel_entry["type"] == REL_OFFICE_DOCUMENT and rel_entry.get("part"):
                    pres_part = rel_entry["part"]
                    break
            if pres_part is None or pres_part not in names:
                raise DiscoveryError("no officeDocument relationship resolves to a real part")
            pres = common.load_zip_xml(zf, pres_part)
            if pres is None:
                raise DiscoveryError("presentation part is not parseable XML")
            spec["presentation_part"] = pres_part
            types = _content_type_overrides(zf)
            spec["presentation_content_type"] = types.get(pres_part)
            pres_rels = _read_rels(zf, pres_part)

            sld_sz = _child(pres, "p:sldSz")
            if sld_sz is not None and sld_sz.get("cx") and sld_sz.get("cy"):
                try:
                    cx, cy = int(sld_sz.get("cx")), int(sld_sz.get("cy"))
                except ValueError:
                    cx = cy = 0
                if cx > 0 and cy > 0:
                    divisor = _gcd(cx, cy)
                    aspect = "%d:%d" % (cx // divisor, cy // divisor)
                    spec["slide_size"] = {
                        "cx": cx, "cy": cy, "type": sld_sz.get("type"),
                        "src": "%s/p:sldSz" % pres_part, "aspect": aspect,
                    }
                    spec["evidence"].append({
                        "claim": "slide_size", "src": pres_part,
                        "xpath": "/p:presentation/p:sldSz[@cx][@cy]",
                    })
            if spec["slide_size"] is None:
                spec["notes"].append("presentation part declares no usable p:sldSz; slide size stays null.")

            master_ids = []
            id_lst = _child(pres, "p:sldMasterIdLst")
            if id_lst is not None:
                for el in id_lst:
                    if common.localname(el.tag) != "sldMasterId":
                        continue
                    master_ids.append(el.get("{%s}id" % R))
            for rid in master_ids:
                rel_entry = pres_rels.get(rid or "")
                part = rel_entry.get("part") if rel_entry else None
                if not part or part not in names:
                    spec["notes"].append("sldMasterId %r does not resolve to a real part" % rid)
                    continue
                master, layout_ids = _parse_master(zf, part)
                if master is None:
                    spec["notes"].append("master %s is not parseable XML" % part)
                    continue
                master_index = {(p["idx"], p["ph_type"]): p for p in master["placeholders"]}
                for _lrid, lpart in layout_ids:
                    if not lpart or lpart not in names:
                        spec["notes"].append("master %s references a missing layout part %r" % (part, lpart))
                        continue
                    layout = _parse_layout(zf, lpart, part, master_index)
                    master["layouts"].append(lpart)
                    spec["layouts"].append(layout)
                spec["masters"].append(master)
                spec["evidence"].append({
                    "claim": "masters", "src": part,
                    "xpath": "/p:sldMaster/p:sldLayoutIdLst",
                })

            for rel_entry in pres_rels.values():
                if rel_entry["type"].endswith("/notesMaster") and rel_entry.get("part"):
                    spec["notes_master"] = rel_entry["part"]
                    break

            theme_parts = []
            for master in spec["masters"]:
                if master.get("theme_rel") and master["theme_rel"] not in theme_parts:
                    theme_parts.append(master["theme_rel"])
            for rel_entry in pres_rels.values():
                if rel_entry["type"].endswith("/theme") and rel_entry.get("part"):
                    if rel_entry["part"] not in theme_parts:
                        theme_parts.append(rel_entry["part"])
            for part in theme_parts:
                theme = _parse_theme(zf, part) if part in names else None
                if theme is not None:
                    spec["theme"].append(theme)
            if spec["theme"]:
                spec["evidence"].append({
                    "claim": "theme_fonts_and_colors", "src": spec["theme"][0]["part"],
                    "xpath": "/a:theme/a:themeElements/a:fontScheme|/a:theme/a:themeElements/a:clrScheme",
                })
            else:
                spec["notes"].append("no theme part could be parsed; theme fonts/colors stay unknown.")

            if spec["slide_size"] is None or not spec["masters"] or not spec["layouts"]:
                spec["notes"].append("template is incomplete for inheritance: slide size, masters or layouts are missing.")
            else:
                spec["ok"] = True
    except (zipfile.BadZipFile, zipfile.LargeZipFile) as exc:
        spec["notes"].append("not a readable OPC package (zip): %s" % exc)
    except (OSError, ET.ParseError, DiscoveryError, ValueError) as exc:
        spec["notes"].append("template could not be inspected: %s: %s" % (type(exc).__name__, exc))
    return spec


def _gcd(a, b):
    while b:
        a, b = b, a % b
    return a or 1


# --------------------------------------------------------------------------------------
# starting deck: per-slide content detail (ADDENDUM v1.3 §A1; everything from real bytes)
# --------------------------------------------------------------------------------------

_TITLE_PH_TYPES = ("title", "ctrTitle")
_BODY_PH_TYPES = ("body", "obj", "subTitle")


def _where_of_ph_type(ph_type):
    """Map a real placeholder type onto the A1 ``where`` vocabulary."""
    if ph_type in _TITLE_PH_TYPES:
        return "title"
    if ph_type in _BODY_PH_TYPES:
        return "body"
    return "other"


def _signature_kind(ph_type):
    """Normalised placeholder kind used for the A1 ``placeholder_signature`` list."""
    if ph_type in _TITLE_PH_TYPES:
        return "title"
    if ph_type in _BODY_PH_TYPES:
        return "body"
    return ph_type


def _iter_sp_tree_children(elem):
    """Yield shape-level children of a ``p:spTree`` in document order, descending groups."""
    if elem is None:
        return
    for child in list(elem):
        local = common.localname(child.tag)
        if local in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"):
            yield child
            if local == "grpSp":
                for nested in _iter_sp_tree_children(child):
                    yield nested


def _paragraph_text_units(container, where):
    """``[{text, where, level}]`` for every non-empty paragraph of one shape/body."""
    units = []
    if container is None:
        return units
    body = container.find("p:txBody", NS)
    if body is None:
        return units
    for para in body.findall("a:p", NS):
        text = "".join((node.text or "") for node in para.iter("{%s}t" % A)).strip()
        if not text:
            continue
        ppr = para.find("a:pPr", NS)
        level = 0
        if ppr is not None and ppr.get("lvl") is not None:
            try:
                level = int(ppr.get("lvl"))
            except (TypeError, ValueError):
                level = 0
        units.append({"text": text, "where": where, "level": level})
    return units


def _shape_text_units(shape):
    """Text units of a real shape, deciding ``where`` from its actual placeholder type."""
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None:
        _idx, ph_type = _ph_key(ph)
        where = _where_of_ph_type(ph_type)
    else:
        where = "other"
    return _paragraph_text_units(shape, where)


def _slide_picture_records(zf, slide_root, slide_part):
    """Real picture bytes referenced by a slide (``rel`` is the media part in the package)."""
    rels = _read_rels(zf, slide_part)
    pictures = []
    for node in slide_root.iter("{%s}pic" % P):
        blip = node.find("p:blipFill/a:blip", NS)
        rid = blip.get("{%s}embed" % R) if blip is not None else None
        rel = rels.get(rid or "")
        target = rel.get("part") if rel else None
        if not target:
            pictures.append({"rel": None, "rid": rid, "sha256": None, "width": None, "height": None,
                             "note": "picture has no resolvable image relationship"})
            continue
        try:
            data = zf.read(target)
        except (KeyError, OSError) as exc:
            pictures.append({"rel": target, "rid": rid, "sha256": None, "width": None, "height": None,
                             "note": "image part could not be read: %s" % exc})
            continue
        kind = common.guess_kind(target)
        size = common.natural_size(data, kind)
        pictures.append({
            "rel": target, "rid": rid, "sha256": common.sha256_bytes(data),
            "width": size[0] if size else None, "height": size[1] if size else None,
            "bytes": len(data),
        })
    return pictures


def _slide_table_records_from_frame(frame):
    """Real table geometry + cell text of one ``p:graphicFrame``.

    Returns ``(records, cell_text_units)`` where ``records`` only carries real counts
    (``a:tr``/``a:tc``); no cell is invented when a table has no ``<a:tbl>``.
    """
    records, cell_units = [], []
    tbl = frame.find("a:graphic/a:graphicData/a:tbl", NS)
    if tbl is None:
        return records, cell_units
    rows = tbl.findall("a:tr", NS)
    cols = 0
    for row in rows:
        cols = max(cols, len(row.findall("a:tc", NS)))
    records.append({"rows": len(rows), "cols": cols})
    for row_number, row in enumerate(rows, 1):
        for cell_number, cell in enumerate(row.findall("a:tc", NS), 1):
            text = "".join((node.text or "") for node in cell.iter("{%s}t" % A)).strip()
            if text:
                cell_units.append({
                    "text": text, "where": "other", "level": 0,
                    "in_table": True,
                    "cell": "%d,%d" % (row_number, cell_number),
                })
    return records, cell_units


def _slide_notes_text(zf, slide_part, notes_cache):
    """Speaker-note text of a slide (empty string when the slide has no notes part)."""
    rels = _read_rels(zf, slide_part)
    for rel in rels.values():
        if not rel["type"].endswith("/notesSlide") or not rel.get("part"):
            continue
        part = rel["part"]
        if part in notes_cache:
            return notes_cache[part]
        root = common.load_zip_xml(zf, part)
        texts = []
        if root is not None:
            for sp in root.iter("{%s}sp" % P):
                for para in _paragraph_text_units(sp, "other"):
                    texts.append(para["text"])
        text = "\n".join(texts).strip()
        notes_cache[part] = text
        return text
    return ""


def _slide_layout_probe(zf, slide_part):
    """``(raw target, resolved part)`` of the slide's real ``slideLayout`` relationship."""
    for rel in _read_rels(zf, slide_part).values():
        if rel["type"].endswith("/slideLayout"):
            return rel.get("target"), rel.get("part")
    return None, None


def start_deck_spec(path, rel, sha):
    """Honest content inventory of an optional starting deck (ADDENDUM v1.3 §A1).

    v1.2 keys (``path/sha256/ok/slides/slide_titles/notes``) are kept byte for byte in
    meaning; v1.3 adds ``slide_details[]`` and the aggregate counters. Every value comes
    from real package bytes; a broken package yields ``ok=false`` plus the reason and
    never raises.
    """
    spec = {
        "path": rel, "sha256": sha, "ok": False,
        "slides": 0, "slide_titles": [],
        "slide_details": [],
        "slide_count": 0, "has_pictures": False, "has_tables": False, "notes_count": 0,
        "notes": [],
    }
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if "[Content_Types].xml" not in names or "_rels/.rels" not in names:
                raise DiscoveryError("package lacks [Content_Types].xml or _rels/.rels")
            bad = zf.testzip()
            if bad:
                raise DiscoveryError("CRC failure in member %s" % bad)
            pres_part = None
            for rel_entry in _read_rels(zf, "").values():
                if rel_entry["type"] == REL_OFFICE_DOCUMENT and rel_entry.get("part"):
                    pres_part = rel_entry["part"]
                    break
            if not pres_part or pres_part not in names:
                raise DiscoveryError("no presentation part")
            pres = common.load_zip_xml(zf, pres_part)
            if pres is None:
                raise DiscoveryError("presentation part is not parseable XML")
            pres_rels = _read_rels(zf, pres_part)
            slide_parts = []
            id_lst = _child(pres, "p:sldIdLst")
            if id_lst is not None:
                for el in id_lst:
                    if common.localname(el.tag) != "sldId":
                        continue
                    rel_entry = pres_rels.get(el.get("{%s}id" % R) or "")
                    if rel_entry and rel_entry.get("part"):
                        slide_parts.append(rel_entry["part"])
            layout_name_cache = {}
            notes_cache = {}
            for index, part in enumerate(slide_parts, 1):
                slide_titles_entry = {"part": part, "title": None}
                detail = {
                    "index": index,
                    "part": part,
                    "layout": None,
                    "layout_probe": None,
                    "layout_name": None,
                    "placeholder_signature": [],
                    "placeholder_types": [],
                    "title": None,
                    "text_units": [],
                    "pictures": [],
                    "notes": "",
                    "tables": [],
                }
                root = common.load_zip_xml(zf, part)
                if root is None:
                    spec["notes"].append("slide part %s is not parseable" % part)
                    detail["parse_note"] = "slide part is not parseable XML"
                    spec["slide_titles"].append(slide_titles_entry)
                    spec["slide_details"].append(detail)
                    continue
                c_sld = _child(root, "p:cSld")
                sp_tree = _child(c_sld, "p:spTree") if c_sld is not None else None
                signature, types = [], []
                for shape in _iter_sp_tree_children(sp_tree):
                    local = common.localname(shape.tag)
                    if local == "sp":
                        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
                        if ph is not None:
                            _idx, ph_type = _ph_key(ph)
                            types.append(ph_type)
                            kind = _signature_kind(ph_type)
                            if kind not in signature:
                                signature.append(kind)
                        for unit in _shape_text_units(shape):
                            detail["text_units"].append(unit)
                    elif local == "pic":
                        # picture placeholders may also carry real text in unusual decks
                        for unit in _shape_text_units(shape):
                            detail["text_units"].append(unit)
                    elif local == "graphicFrame":
                        tables, cell_units = _slide_table_records_from_frame(shape)
                        for table in tables:
                            detail["tables"].append(dict(table, index=len(detail["tables"]) + 1))
                        table_index = len(detail["tables"])
                        for unit in cell_units:
                            unit["table"] = table_index
                            detail["text_units"].append(unit)
                title = None
                for sp, ph, _cnv in _placeholder_shapes(root):
                    if _ph_key(ph)[1] in _TITLE_PH_TYPES:
                        title = "".join((node.text or "") for node in sp.iter("{%s}t" % A)).strip()
                        if title:
                            break
                detail["title"] = title or None
                slide_titles_entry["title"] = title or None
                detail["placeholder_signature"] = signature
                detail["placeholder_types"] = types
                raw_target, layout_part = _slide_layout_probe(zf, part)
                detail["layout_probe"] = raw_target
                detail["layout"] = layout_part
                if layout_part and layout_part in names:
                    if layout_part not in layout_name_cache:
                        layout_root = common.load_zip_xml(zf, layout_part)
                        c_sld_layout = _child(layout_root, "p:cSld") if layout_root is not None else None
                        layout_name_cache[layout_part] = _attr(c_sld_layout, "name")
                    detail["layout_name"] = layout_name_cache[layout_part]
                elif layout_part:
                    spec["notes"].append("slide %d points at a layout part missing from the package: %s"
                                         % (index, layout_part))
                detail["pictures"] = _slide_picture_records(zf, root, part)
                detail["notes"] = _slide_notes_text(zf, part, notes_cache)
                spec["slide_titles"].append(slide_titles_entry)
                spec["slide_details"].append(detail)
            spec["slides"] = len(slide_parts)
            spec["slide_count"] = len(slide_parts)
            spec["has_pictures"] = any(item["pictures"] for item in spec["slide_details"])
            spec["has_tables"] = any(item["tables"] for item in spec["slide_details"])
            spec["notes_count"] = sum(1 for item in spec["slide_details"] if item.get("notes"))
            spec["ok"] = spec["slides"] > 0
    except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError, ET.ParseError, DiscoveryError) as exc:
        spec["notes"].append("start deck could not be inspected: %s: %s" % (type(exc).__name__, exc))
    return spec


# --------------------------------------------------------------------------------------
# source parsing (markdown / text / docx / pdf)
# --------------------------------------------------------------------------------------

def _strip_markup(text):
    """Presentation markup removed for deck use (``**bold**``, ``#``/``>`` prefixes).

    ``unit["text"]`` / ``caption`` stay byte-exact so the audit can compare against the
    source; ``text_plain`` / ``caption_plain`` are what a slide should actually show.
    """
    out = (text or "").replace("**", "").replace("__", "")
    out = re.sub(r"(?m)^\s{0,3}#{1,6}\s*", "", out)
    out = re.sub(r"(?m)^\s{0,3}>\s?", "", out)
    return re.sub(r"\s*[\r\n]+\s*", " ", out).strip()


def _unit(uid, kind, text, locator, section=None, ordinal=0, extra=None):
    text = (text or "").strip()
    rec = {
        "id": uid,
        "kind": kind,
        "text": text,
        "text_plain": _strip_markup(text),
        "locator": locator,
        "section": section,
        "ordinal": ordinal,
        "sha256": common.sha256_bytes(text.encode("utf-8")),
    }
    if extra:
        rec.update(extra)
    return rec


def _classify_block(block, in_abstract):
    first = block.splitlines()[0].strip()
    if first.startswith(">"):
        return "abstract" if in_abstract else "text"
    if in_abstract:
        return "abstract"
    return "text"


# --------------------------------------------------------------------------------------
# v1.3: real table cell data (ADDENDUM §A6 table requirement; never invented)
# --------------------------------------------------------------------------------------

_MD_TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _split_md_table_row(line):
    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|"):
        text = text[:-1]
    return [cell.strip() for cell in text.split("|")]


def extract_markdown_tables(lines, section_of_line):
    """Every real markdown pipe table: header + body rows, with its source line."""
    tables = []
    index = 0
    while index < len(lines) - 1:
        current = lines[index].strip()
        following = lines[index + 1].strip()
        if current.startswith("|") and following.startswith("|") and _MD_TABLE_SEP_RE.match(following):
            header = _split_md_table_row(current)
            rows = []
            cursor = index + 2
            while cursor < len(lines) and lines[cursor].strip().startswith("|"):
                rows.append(_split_md_table_row(lines[cursor]))
                cursor += 1
            section = section_of_line[index] if index < len(section_of_line) else None
            tables.append({
                "kind": "markdown pipe table",
                "line": index + 1,
                "section": section["id"] if section else None,
                "section_title": section["title"] if section else None,
                "header": header, "rows": rows,
            })
            index = cursor
            continue
        index += 1
    return tables


def _docx_table_records(root):
    """Every real ``w:tbl`` in document order: first row as header, remaining rows as data."""
    out = []
    body = root.find("{%s}body" % W_NS)
    scope = body if body is not None else root
    for order, element in enumerate(
            [el for el in scope.iter() if common.localname(el.tag) == "tbl"], 1):
        rows = []
        for row in element.findall("{%s}tr" % W_NS):
            cells = [_docx_paragraph_text(cell).strip() for cell in row.findall("{%s}tc" % W_NS)]
            if cells:
                rows.append(cells)
        if not rows:
            continue
        out.append({"kind": "docx w:tbl", "order": order,
                    "header": rows[0], "rows": rows[1:]})
    return out


def _table_locator_line(entry):
    match = re.search(r"line (\d+)", entry.get("locator") or "")
    return int(match.group(1)) if match else None


def _assign_table_cells(entry, table, link):
    entry["header"] = list(table.get("header") or [])
    entry["rows"] = [list(row) for row in (table.get("rows") or [])]
    entry["data_source"] = table.get("kind")
    entry["data_link"] = link
    entry["data_locator"] = ("line %d" % table["line"]) if table.get("line") else None
    widths = [len(entry["header"])] + [len(row) for row in entry["rows"]]
    cols = max(widths) if widths else 0
    entry["shape"] = {"header_rows": 1 if entry["header"] else 0,
                      "data_rows": len(entry["rows"]), "cols": cols}


def _attach_table_data(table_entries, parsed_tables):
    """Link real parsed cells to ``tables[]`` entries; never invent a row.

    Priority: a section heading that names the table (e.g. ``## 表1 消融实验汇总``) →
    nearest table in the caption's own section → document order when counts match 1:1.
    Anything unlinked keeps ``header=[]``/``rows=[]`` plus a ``data_note`` reason.
    """
    for entry in table_entries:
        entry.setdefault("header", [])
        entry.setdefault("rows", [])
        entry.setdefault("data_source", None)
        entry.setdefault("data_note", None)
    if not table_entries:
        return
    if not parsed_tables:
        for entry in table_entries:
            if not entry.get("data_note"):
                entry["data_note"] = ("no markdown/docx table with real cells was found in this "
                                      "source; header/rows stay empty")
        return
    assigned = set()
    for entry in table_entries:  # pass 1: section heading names this table
        label_plain = entry["label"].replace(" ", "")
        for index, table in enumerate(parsed_tables):
            if index in assigned:
                continue
            title = (table.get("section_title") or "").replace(" ", "")
            if label_plain and label_plain in title:
                _assign_table_cells(entry, table, "section heading names this table: %r" % table.get("section_title"))
                assigned.add(index)
                break
    for entry in table_entries:  # pass 2: nearest table in the same section
        if entry.get("data_source"):
            continue
        line = _table_locator_line(entry)
        candidates = []
        for index, table in enumerate(parsed_tables):
            if index in assigned or not table.get("line"):
                continue
            if table.get("section") != entry.get("section"):
                continue
            distance = abs((table["line"] or 0) - (line if line is not None else table["line"]))
            candidates.append((distance, index))
        if candidates:
            distance, index = min(candidates)
            _assign_table_cells(entry, parsed_tables[index],
                                "nearest table in the same section (line distance %d)" % distance)
            assigned.add(index)
    free = [index for index in range(len(parsed_tables)) if index not in assigned]
    unassigned = [entry for entry in table_entries if not entry.get("data_source")]
    if unassigned and len(parsed_tables) == len(table_entries) and len(free) == len(unassigned):
        for entry, index in zip(unassigned, free):  # pass 3: one caption per table, document order
            _assign_table_cells(entry, parsed_tables[index],
                                "assigned in document order (one real table per caption)")
            assigned.add(index)
    for entry in table_entries:  # honest reason when nothing could be linked
        if not entry.get("data_source") and not entry.get("data_note"):
            entry["data_note"] = ("%d real table(s) exist in this source but no unambiguous link to %s "
                                  "was found; header/rows stay empty" % (len(parsed_tables), entry["label"]))


def parse_markdown(text, rel):
    """Section-aware extraction with ``line N`` locators.

    A paragraph block is flushed at blank lines and headings only, and is always kept as a
    ``text``/``abstract`` unit. Figure/table captions are extracted *out of* the block as
    the single sentence that carries each number, so ``图1``/``图2``/``表1`` mentioned in
    one result paragraph produce three independent entries and no body text is lost.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    sections, units, figures, tables = [], [], [], []
    section_of_line = []
    current = None
    ordinal = 0
    for line in lines:
        match = _HEADING_RE.match(line.strip())
        if match:
            ordinal += 1
            current = {
                "id": "S%d" % ordinal,
                "level": len(match.group(1)),
                "title": match.group(2).strip(),
                "ordinal": ordinal,
            }
            sections.append(current)
        section_of_line.append(current)

    uid = 0

    def new_unit(kind, body, line_number, section):
        nonlocal uid
        uid += 1
        return _unit("u%04d" % uid, kind, body, "line %d" % line_number,
                     section["id"] if section else None, uid)

    def in_abstract(section):
        return bool(section and re.search(r"摘要|abstract", section["title"], re.IGNORECASE))

    def collect_captions(block, start_line, section):
        for entry in extract_caption_entries(block):
            line_number = start_line + block[: entry["offset"]].count("\n")
            record = new_unit("figure_caption" if entry["kind"] == "figure" else "table_caption",
                              entry["caption"], line_number, section)
            units.append(record)
            _record_caption(entry, record, "line %d" % line_number,
                            "%s#L%d" % (rel, line_number), figures, tables)

    buffer_lines = []
    buffer_start = None
    buffer_section = None
    buffer_quote = False

    def flush():
        nonlocal buffer_lines, buffer_start, buffer_section, buffer_quote
        if not buffer_lines:
            return
        block = "\n".join(buffer_lines).strip()
        if block:
            kind = "abstract" if in_abstract(buffer_section) else "text"
            record = new_unit(kind, block, buffer_start + 1, buffer_section)
            if (len(buffer_lines) >= 2
                    and all(line.strip().startswith("|") for line in buffer_lines)
                    and any(_MD_TABLE_SEP_RE.match(line.strip()) for line in buffer_lines)):
                # real markdown pipe table: cell data lives in tables[]; flag the text unit so
                # a planner does not also render the raw pipe rows as bullets.
                record["in_table"] = True
                record["table_data_source"] = "markdown pipe table starting at line %d" % (buffer_start + 1)
            units.append(record)
            collect_captions(block, buffer_start + 1, buffer_section)
        buffer_lines = []
        buffer_start = None
        buffer_section = None
        buffer_quote = False

    number = 0
    while number < len(lines):
        raw = lines[number]
        stripped = raw.strip()
        section = section_of_line[number]
        if not stripped:
            flush()
            number += 1
            continue
        heading = _HEADING_RE.match(stripped)
        if heading:
            flush()
            units.append(new_unit("heading", stripped, number + 1, section))
            number += 1
            continue
        if stripped.startswith(">"):
            if buffer_lines and not buffer_quote:
                flush()
            if not buffer_lines:
                buffer_start = number
                buffer_section = section
                buffer_quote = True
            buffer_lines.append(stripped)
            number += 1
            continue
        if buffer_lines and (buffer_quote or buffer_section is not section):
            flush()
        if not buffer_lines:
            buffer_start = number
            buffer_section = section
        buffer_lines.append(stripped)
        number += 1
    flush()
    parsed_tables = extract_markdown_tables(lines, section_of_line)
    _attach_table_data(tables, parsed_tables)
    return sections, units, figures, tables


def _docx_paragraph_text(paragraph):
    chunks = []
    for node in paragraph.iter():
        local = common.localname(node.tag)
        if local == "t" and node.text:
            chunks.append(node.text)
        elif local == "tab":
            chunks.append("\t")
        elif local in ("br", "cr"):
            chunks.append("\n")
    return "".join(chunks)


def parse_docx(path, rel):
    sections, units, figures, tables = [], [], [], []
    try:
        with zipfile.ZipFile(path) as zf:
            if "word/document.xml" not in zf.namelist():
                return sections, units, figures, tables, {
                    "tool": "docx-xml", "status": "failed",
                    "detail": "word/document.xml missing from the docx package",
                }
            root = ET.fromstring(zf.read("word/document.xml"))
    except (zipfile.BadZipFile, OSError, ET.ParseError) as exc:
        return sections, units, figures, tables, {
            "tool": "docx-xml", "status": "failed",
            "detail": "%s: %s" % (type(exc).__name__, exc),
        }
    ordinal = 0
    uid = 0
    current = None
    for index, para in enumerate([p for p in root.iter() if common.localname(p.tag) == "p"], 1):
        text = _docx_paragraph_text(para).strip()
        if not text:
            continue
        ppr = para.find("{%s}pPr" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        style = None
        if ppr is not None:
            style_el = ppr.find("{%s}pStyle" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            style = style_el.get("{%s}val" % "http://schemas.openxmlformats.org/wordprocessingml/2006/main") if style_el is not None else None
        is_heading = bool(style and re.search(r"heading|标题|title", style, re.IGNORECASE))
        locator = "word/document.xml#p%d" % index
        if is_heading:
            ordinal += 1
            level_match = re.search(r"(\d)", style or "")
            current = {
                "id": "S%d" % ordinal,
                "level": int(level_match.group(1)) if level_match else 1,
                "title": text,
                "ordinal": ordinal,
            }
            sections.append(current)
            uid += 1
            units.append(_unit("u%04d" % uid, "heading", text, locator, current["id"], uid))
            continue
        in_abstract = bool(current and re.search(r"摘要|abstract", current["title"], re.IGNORECASE))
        kind = _classify_block(text, in_abstract)
        uid += 1
        units.append(_unit("u%04d" % uid, kind, text, locator, current["id"] if current else None, uid))
        for entry in extract_caption_entries(text):
            uid += 1
            record = _unit("u%04d" % uid,
                           "figure_caption" if entry["kind"] == "figure" else "table_caption",
                           entry["caption"], locator, current["id"] if current else None, uid)
            units.append(record)
            _record_caption(entry, record, locator, locator, figures, tables)
    _attach_table_data(tables, _docx_table_records(root))
    return sections, units, figures, tables, {
        "tool": "docx-xml", "status": "ok", "detail": "%d paragraphs with text" % len(units),
    }


PDF_TEXT_CLI_NAMES = ("pdf-text-tool", "pdftext", "ds-pdf-text")
PDF_TEXT_ENV_VAR = "PDF_TEXT_TOOL"
PDF_PY_MODULES = ("pypdf", "PyPDF2", "pdfplumber", "fitz")

PDF_TEXT_DETAIL = ("no PDF text extractor succeeded; source_inventory.units stays empty and no "
                   "body text is invented")


def _pdf_tool_chain():
    """Ordered list of ``(tool, path)`` candidates available on this machine (A6 order)."""
    chain = []
    exe = shutil.which("pdftotext")
    if exe:
        chain.append(("pdftotext", exe))
    mutool = shutil.which("mutool")
    if mutool:
        chain.append(("mutool", mutool))
    for module in PDF_PY_MODULES:
        try:
            __import__(module)
        except Exception:
            continue
        chain.append((module, None))
    env_tool = os.environ.get(PDF_TEXT_ENV_VAR)
    if env_tool:
        chain.append(("$" + PDF_TEXT_ENV_VAR, env_tool))
    for name in PDF_TEXT_CLI_NAMES:
        exe = shutil.which(name)
        if exe:
            chain.append((name, exe))
    return chain


def _detect_pdf_tool():
    """First usable candidate in the A6 chain; ``(None, None)`` when nothing is available."""
    chain = _pdf_tool_chain()
    return chain[0] if chain else (None, None)


def _pdf_attempt(attempts, tool, status, detail):
    """Record one real attempt; ``status ∈ {ok, failed, unavailable}`` (never invented)."""
    attempts.append({"tool": tool, "status": status, "detail": detail})


def _clean_tool_stderr(text):
    """Last real stderr line with the local home directory masked (no author paths leak)."""
    text = text or ""
    home = os.path.expanduser("~")
    if home:
        text = text.replace(home, "~")
        text = text.replace(home.replace("\\", "/"), "~")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return lines[-1][:300] if lines else ""


def _run_pdf_command(argv, timeout=120):
    """Run one external extractor; returns ``(text|None, detail)`` without raising."""
    try:
        proc = subprocess.run(argv, capture_output=True, timeout=timeout)
    except (OSError, subprocess.SubprocessError) as exc:
        return None, "%s: %s" % (type(exc).__name__, exc)
    if proc.returncode != 0:
        stderr = _clean_tool_stderr(proc.stderr.decode("utf-8", "replace"))
        return None, "exit %d%s" % (proc.returncode, (": " + stderr) if stderr else "")
    # a zero exit means the text stream is the payload; stderr here is tool logging noise
    return proc.stdout.decode("utf-8", "replace"), ""


def _extract_pdf_with_module(module, path):
    """Text of a PDF through one real python library; raises nothing unexpected to callers."""
    if module in ("pypdf", "PyPDF2"):
        lib = __import__(module)
        reader = lib.PdfReader(path)
        pages = [(page.extract_text() or "") for page in reader.pages]
        return "\f".join(pages), "%d pages" % len(pages)
    if module == "pdfplumber":
        import pdfplumber  # type: ignore
        with pdfplumber.open(path) as pdf:
            pages = [(page.extract_text() or "") for page in pdf.pages]
        return "\f".join(pages), "%d pages" % len(pages)
    if module == "fitz":
        import fitz  # type: ignore
        doc = fitz.open(path)
        pages = [(page.get_text() or "") for page in doc]
        doc.close()
        return "\f".join(pages), "%d pages" % len(pages)
    raise ValueError("unknown module %r" % module)


def extract_pdf(path, rel):
    """Probe the A6 extractor chain, recording **every** attempt in ``extraction.attempts``.

    Returns ``(text, extraction)``; ``text`` is ``None`` only when every candidate failed,
    in which case ``status=failed`` and no body text is fabricated by the caller.
    """
    attempts = []

    def succeed(tool, text, detail):
        _pdf_attempt(attempts, tool, "ok", detail or "text extracted")
        return text, {"tool": tool, "status": "ok", "detail": detail or "text extracted",
                      "attempts": attempts}

    def failed(tool, detail):
        _pdf_attempt(attempts, tool, "failed", detail)
        return False

    def unavailable(tool, detail):
        _pdf_attempt(attempts, tool, "unavailable", detail)

    # 1) pdftotext -layout FIRST: it preserves the blank paragraph lines that let a real PDF
    #    become several source units (a plain pdftotext run flattens a page into one block).
    pdftotext = shutil.which("pdftotext")
    if not pdftotext:
        unavailable("pdftotext", "not found on PATH")
    else:
        text, detail = _run_pdf_command([pdftotext, "-layout", path, "-"])
        if text is not None and text.strip():
            return succeed("pdftotext -layout", text, detail or "layout-preserving text stream")
        failed("pdftotext -layout", detail or "produced no text")
        text, detail = _run_pdf_command([pdftotext, path, "-"])
        if text is not None and text.strip():
            return succeed("pdftotext", text, detail or "plain text stream")
        failed("pdftotext", detail or "produced no text")
    mutool = shutil.which("mutool")
    if not mutool:
        unavailable("mutool", "not found on PATH")
    else:
        text, detail = _run_pdf_command([mutool, "draw", "-F", "txt", "-o", "-", path])
        if text is not None and text.strip():
            return succeed("mutool draw -F txt", text, detail or "mutool text device")
        failed("mutool draw -F txt", detail or "produced no text")
    if pdftotext:
        # legacy OEM code-page variant declared by the task statement (CJK/dos encodings)
        text, detail = _run_pdf_command([pdftotext, "-layout", "-enc", "OEM", path, "-"])
        if text is not None and text.strip():
            return succeed("pdftotext -layout -enc OEM", text, detail or "OEM-encoded text stream")
        failed("pdftotext -layout -enc OEM", detail or "produced no text")

    # 3) python libraries (probe order fixed by A6; each one is genuinely attempted)
    for module in PDF_PY_MODULES:
        try:
            __import__(module)
        except Exception as exc:
            unavailable(module, "import failed: %s: %s" % (type(exc).__name__, exc))
            continue
        try:
            text, detail = _extract_pdf_with_module(module, path)
        except Exception as exc:
            failed(module, "%s: %s" % (type(exc).__name__, exc))
            continue
        if text and text.strip():
            return succeed(module, text, detail)
        failed(module, detail or "extractor returned no text")

    # 4) the task-declared CLI tools (env var first, then the documented PATH names)
    cli_candidates = []
    env_tool = os.environ.get(PDF_TEXT_ENV_VAR)
    if env_tool:
        cli_candidates.append(("$" + PDF_TEXT_ENV_VAR, env_tool))
    for name in PDF_TEXT_CLI_NAMES:
        exe = shutil.which(name)
        if exe:
            cli_candidates.append((name, exe))
    if not cli_candidates:
        for name in PDF_TEXT_CLI_NAMES:
            unavailable(name, "not found on PATH and %s is unset" % PDF_TEXT_ENV_VAR)
    for label, exe in cli_candidates:
        text, detail = _run_pdf_command([exe, path])
        if text is not None and text.strip():
            return succeed(label, text, detail or "external pdf text tool")
        failed(label, detail or "produced no text")

    return None, {
        "tool": "none", "status": "failed",
        "detail": "%s (%d attempts recorded)" % (PDF_TEXT_DETAIL, len(attempts)),
        "attempts": attempts,
    }


PDF_UNIT_TARGET_CHARS = 300
_PDF_END_PUNCT = "。．.，,；;：:"
_PDF_TITLE_MAX_CHARS = 40
_PDF_HEADING_MAX_CHARS = 60
# Real PDF headings carry no markdown '#': 一、/（一）/1./1.1/1)/第X章 + ALL-CAPS short lines.
_PDF_HEADING_RE = re.compile(
    r"^(?:"
    r"第?\s*[一二三四五六七八九十百]+\s*[、.．)）]"          # 一、 二. 三)
    r"|第\s*[一二三四五六七八九十百\d]+\s*[章节部分]"        # 第一章 / 第 2 部分
    r"|[（(]\s*[一二三四五六七八九十\d]+\s*[）)]"            # （一） (2)
    r"|\d{1,2}(?:\.\d{1,2}){0,3}\s*[、.．)）]"               # 1. 1.1. 2)
    r"|\d{1,2}(?:\.\d{1,2})+\s+\S"                           # 1.1 方法
    r"|[A-Z][A-Z0-9 \-_/&.,]{2,60}$"                          # RESULTS / METHOD A
    r")"
)
_PDF_SENTENCE_BREAK_RE = re.compile(r"(?<=[。！？!?])\s*|(?<=\.)\s+(?=[A-Z0-9(])")


def _pdf_normalise_page(text):
    return (text or "").replace("\r\n", "\n").replace("\r", "\n")


def _pdf_page_blocks(lines):
    """``[(start_line_index, [lines])]`` — blank lines separate real paragraphs."""
    blocks = []
    index = 0
    while index < len(lines):
        if not lines[index].strip():
            index += 1
            continue
        start = index
        chunk = []
        while index < len(lines) and lines[index].strip():
            chunk.append(lines[index].rstrip())
            index += 1
        blocks.append((start, chunk))
    return blocks


def _pdf_is_heading(text):
    text = (text or "").strip()
    if not text or len(text) > _PDF_HEADING_MAX_CHARS:
        return False
    return bool(_PDF_HEADING_RE.match(text))


def _pdf_title_candidate(text, next_text):
    """First PDF line as the document title: short, no final punctuation, next line longer."""
    text = (text or "").strip()
    if not text or len(text) > _PDF_TITLE_MAX_CHARS:
        return False
    if text[-1] in _PDF_END_PUNCT:
        return False
    return len((next_text or "").strip()) > len(text)


def _pdf_block_segments(start, chunk):
    """Split one paragraph block into ``(line_number, text, single_line)`` segments.

    A multi-line block whose first line is a real heading becomes a heading segment plus a
    remainder paragraph, so the heading text is never swallowed into body prose.
    """
    joined = " ".join(part.strip() for part in chunk).strip()
    if not joined:
        return []
    if len(chunk) > 1 and _pdf_is_heading(chunk[0].strip()):
        segments = [(start + 1, chunk[0].strip(), True)]
        rest = " ".join(part.strip() for part in chunk[1:]).strip()
        if rest:
            segments.append((start + 2, rest, False))
        return segments
    return [(start + 1, joined, len(chunk) == 1)]


def _split_long_pdf_text(text, limit=PDF_UNIT_TARGET_CHARS):
    """Split a long real paragraph on sentence boundaries (hard split only as a last resort)."""
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]
    parts = [part.strip() for part in _PDF_SENTENCE_BREAK_RE.split(text) if part and part.strip()]
    chunks, buffer = [], ""
    for part in parts:
        if len(part) > limit:
            if buffer:
                chunks.append(buffer)
                buffer = ""
            while len(part) > limit:
                chunks.append(part[:limit])
                part = part[limit:]
            buffer = part
            continue
        if not buffer:
            buffer = part
        elif len(buffer) + len(part) <= limit:
            separator = " " if (buffer[-1].isascii() and part[0].isascii()) else ""
            buffer += separator + part
        else:
            chunks.append(buffer)
            buffer = part
    if buffer:
        chunks.append(buffer)
    return [chunk.strip() for chunk in chunks if chunk.strip()]


def parse_pdf(text, rel):
    """Page-aware PDF extraction: paragraphs, then line-level fallback, then real headings.

    ``pdftotext -layout`` preserves paragraph blank lines, so a normal PDF yields one unit
    per paragraph.  When a PDF has no blank lines at all, the whole document falls back to
    one unit per real non-empty line (never a single invented block).  Headings/title are
    recognised from the text itself (no markdown); over-long paragraphs are split on real
    sentence boundaries.  Locators stay ``p.N line M`` and only extracted characters are
    ever emitted.
    """
    pages = [_pdf_normalise_page(page) for page in (text or "").split("\f")]
    page_blocks = []
    nonempty_lines = 0
    for page_number, page in enumerate(pages, 1):
        lines = page.split("\n")
        nonempty_lines += sum(1 for line in lines if line.strip())
        for start, chunk in _pdf_page_blocks(lines):
            page_blocks.append((page_number, start, chunk))
    if len(page_blocks) < 3 and nonempty_lines >= 5:
        # no paragraph blank lines in this PDF: fall back to one block per real line
        line_blocks = []
        for page_number, page in enumerate(pages, 1):
            for index, line in enumerate(page.split("\n")):
                if line.strip():
                    line_blocks.append((page_number, index, [line]))
        page_blocks = line_blocks

    segments = []
    for page_number, start, chunk in page_blocks:
        for line_number, segment_text, single_line in _pdf_block_segments(start, chunk):
            segments.append((page_number, line_number, segment_text, single_line))

    sections, units, figures, tables = [], [], [], []
    ordinal = 0
    uid = 0
    current = None
    first_segment = True
    for position, (page_number, line_number, block, single_line) in enumerate(segments):
        locator = "p.%d line %d" % (page_number, line_number)
        next_text = segments[position + 1][2] if position + 1 < len(segments) else ""
        is_first = first_segment
        first_segment = False
        markdown_heading = _HEADING_RE.match(block)
        if markdown_heading:
            ordinal += 1
            current = {"id": "S%d" % ordinal, "level": len(markdown_heading.group(1)),
                       "title": markdown_heading.group(2).strip(), "ordinal": ordinal}
            sections.append(current)
            uid += 1
            units.append(_unit("u%04d" % uid, "heading", block, locator, current["id"], uid))
            continue
        if is_first and _pdf_title_candidate(block, next_text):
            ordinal += 1
            current = {"id": "S%d" % ordinal, "level": 1, "title": block, "ordinal": ordinal,
                       "title_source": "first PDF line (<=40 chars, no final punctuation, next line longer)"}
            sections.append(current)
            uid += 1
            record = _unit("u%04d" % uid, "heading", block, locator, current["id"], uid)
            record["is_title"] = True
            units.append(record)
            continue
        if single_line and _pdf_is_heading(block):
            ordinal += 1
            current = {"id": "S%d" % ordinal, "level": 1, "title": block, "ordinal": ordinal,
                       "title_source": "PDF line matches a heading pattern (一、/1./1.1/（一）/ALL CAPS)"}
            sections.append(current)
            uid += 1
            units.append(_unit("u%04d" % uid, "heading", block, locator, current["id"], uid))
            continue
        in_abstract = bool(current and re.search(r"摘要|abstract", current["title"], re.IGNORECASE))
        kind = _classify_block(block, in_abstract)
        for piece in _split_long_pdf_text(block):
            uid += 1
            units.append(_unit("u%04d" % uid, kind, piece, locator,
                               current["id"] if current else None, uid))
        for entry in extract_caption_entries(block):
            uid += 1
            record = _unit("u%04d" % uid,
                           "figure_caption" if entry["kind"] == "figure" else "table_caption",
                           entry["caption"], locator, current["id"] if current else None, uid)
            units.append(record)
            _record_caption(entry, record, locator, "%s#%s" % (rel, locator), figures, tables)

    if sections:
        # text before the first detected heading/title belongs to the first real section
        first_id = sections[0]["id"]
        for unit in units:
            if unit.get("section") is None:
                unit["section"] = first_id
    elif units:
        # no explicit title/heading line: keep the real text renderable without inventing one
        sections.append({"id": "S1", "level": 1, "title": "", "ordinal": 1,
                         "title_source": "no explicit heading/title line detected in the PDF"})
        for unit in units:
            if unit.get("section") is None:
                unit["section"] = "S1"
    _attach_table_data(tables, [])
    return sections, units, figures, tables


def parse_source(path, rel, kind):
    """Dispatch by real kind; returns ``(sections, units, figures, tables, extraction)``."""
    if kind == "docx":
        return parse_docx(path, rel)
    if kind == "pdf":
        text, extraction = extract_pdf(path, rel)
        if text is None:
            return [], [], [], [], extraction
        sections, units, figures, tables = parse_pdf(text, rel)
        return sections, units, figures, tables, extraction
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
            text = handle.read()
    except OSError as exc:
        return [], [], [], [], {"tool": "plain", "status": "failed",
                                "detail": "%s: %s" % (type(exc).__name__, exc)}
    sections, units, figures, tables = parse_markdown(text, rel)
    return sections, units, figures, tables, {
        "tool": "plain", "status": "ok", "detail": "%d characters read" % len(text),
    }


# --------------------------------------------------------------------------------------
# case.json / contract parsing
# --------------------------------------------------------------------------------------

def _first_str(mapping, keys):
    for key in keys:
        if isinstance(mapping, dict) and isinstance(mapping.get(key), str) and mapping[key].strip():
            return mapping[key].strip()
    return None


def _contract_from_task_md(text):
    for block in re.findall(r"```(?:json)?\s*(.*?)```", text or "", re.DOTALL):
        try:
            data = json.loads(block)
        except ValueError:
            continue
        if isinstance(data, dict) and isinstance(data.get("output_contract"), (dict, str, list)):
            return data
    return None


def _find_output_contract(case):
    for key in ("output_contract", "outputContract", "output_template", "result_contract", "contract", "output"):
        value = case.get(key)
        if isinstance(value, str):
            try:
                value = json.loads(value)
            except ValueError:
                continue
        if isinstance(value, (dict, list)) and value:
            return value, "case.json.%s" % key
    for wrapper in ("identification", "case", "task", "output"):
        nested = case.get(wrapper)
        if isinstance(nested, dict):
            found = _find_output_contract(nested)
            if found[0]:
                return found[0], "case.json.%s" % wrapper
    return None, None


def _dotted_leaves(value, prefix="", out=None):
    out = [] if out is None else out
    if isinstance(value, dict) and value:
        for key, item in value.items():
            _dotted_leaves(item, ("%s.%s" % (prefix, key)) if prefix else str(key), out)
    elif isinstance(value, list) and value:
        for item in value:
            _dotted_leaves(item, prefix, out)
    elif prefix:
        out.append(prefix)
    return out


def _materials_of(case):
    for key in ("materials", "inputs", "files", "resources", "assets", "attachments"):
        value = case.get(key)
        if isinstance(value, list) and value:
            return value, key
    return [], None


def _material_rel(entry):
    if isinstance(entry, str):
        return entry
    if not isinstance(entry, dict):
        return None
    return _first_str(entry, ("rel", "path", "file", "filename", "name", "file_path", "rel_path", "target"))


def _normalise_rel(value):
    value = (value or "").strip().strip("`\"'").replace("\\", "/")
    value = re.sub(r"^\./", "", value)
    if value.startswith("/input/"):
        value = value[len("/input/"):]
    elif value.startswith("input/"):
        value = value[len("input/"):]
    return posixpath.normpath(value) if value else value


def classify_roles(case, inputs, notes):
    """Confirm roles only from explicit, unique evidence. Conflicts stay unconfirmed."""
    by_rel = {item["rel"]: item for item in inputs}
    hints = {}
    materials, key = _materials_of(case)
    for index, entry in enumerate(materials):
        rel = _normalise_rel(_material_rel(entry))
        if not rel:
            notes.append("case.json.%s[%d] has no usable path key; ignored." % (key or "materials", index))
            continue
        if rel not in by_rel:
            notes.append("case.json.%s[%d] points at %r, which is not present under the input root." % (key or "materials", index, rel))
            continue
        if not isinstance(entry, dict):
            continue
        hint = _first_str(entry, ("role_hint", "role", "kind_hint", "purpose", "roleHint"))
        if not hint:
            continue
        lowered = hint.lower()
        matched = [role for role, words in ROLE_ALIASES.items() if any(word in lowered for word in words)]
        for role in matched:
            hints.setdefault(role, []).append((rel, "case.json.materials[%d].role_hint=%r" % (index, hint)))
        if not matched:
            notes.append("unrecognised role_hint %r for %s; kept unconfirmed." % (hint, rel))
        figure_hint = _first_str(entry, ("figure_hint", "label", "number"))
        if figure_hint:
            by_rel[rel]["figure_hint"] = figure_hint

    detected = {"template": [], "source": [], "start_deck": [], "figures": []}
    for role in detected:
        rels = []
        for rel, _reason in hints.get(role, []):
            if rel not in rels:
                rels.append(rel)
        if role == "figures":
            # several explicit figure materials are the normal case, not a conflict
            detected["figures"] = rels
        elif len(rels) == 1:
            detected[role] = rels
        elif len(rels) > 1:
            notes.append("role %r has %d explicit candidates (%s); left unconfirmed." % (role, len(rels), ", ".join(rels)))

    # Unique-by-kind fallback, recorded as evidence rather than silently assumed.
    if not detected["template"]:
        candidates = [item["rel"] for item in inputs if item["kind"] in PPTX_KINDS]
        if len(candidates) == 1:
            detected["template"] = candidates
            notes.append("template confirmed by unique pptx-like input kind: %s" % candidates[0])
        elif len(candidates) > 1:
            notes.append("multiple pptx-like inputs and no explicit template role: %s; template unconfirmed." % ", ".join(candidates))
    if not detected["source"]:
        candidates = [item["rel"] for item in inputs if item["kind"] in ("md", "txt", "pdf", "docx", "tex")]
        if len(candidates) == 1:
            detected["source"] = candidates
            notes.append("source confirmed by unique text-like input kind: %s" % candidates[0])
        elif len(candidates) > 1:
            notes.append("multiple text-like inputs and no explicit source role: %s; source unconfirmed." % ", ".join(candidates))
    images = [item["rel"] for item in inputs if item["kind"] in IMAGE_KINDS]
    for rel in images:
        if rel not in detected["figures"]:
            detected["figures"].append(rel)
    for role in ("template", "source", "start_deck"):
        for rel in detected[role]:
            item = by_rel.get(rel)
            if item and role == "template" and item["kind"] not in PPTX_KINDS:
                notes.append("template candidate %s has kind %s, which is not an editable OOXML deck." % (rel, item["kind"]))
    for rel in detected["start_deck"]:
        item = by_rel.get(rel)
        if item and item["kind"] not in PPTX_KINDS:
            notes.append("start_deck candidate %s has kind %s, which cannot be opened as a deck." % (rel, item["kind"]))
    return detected, by_rel


def _attach_figure_assets(figures, inputs, detected):
    """Attach real image bytes to captions: explicit figure_hint first, then name mention, then order."""
    images = [item for item in inputs if item["kind"] in IMAGE_KINDS]
    by_rel = {item["rel"]: item for item in inputs}
    hint_map = {}
    for role_rel in detected.get("figures", []):
        item = by_rel.get(role_rel)
        if item and item.get("figure_hint"):
            hint_map[_normalise_rel(item["figure_hint"])] = role_rel
    for index, figure in enumerate(figures):
        label_key = _normalise_rel(figure["label"])
        asset = None
        origin = None
        if label_key in hint_map:
            asset, origin = hint_map[label_key], "explicit figure_hint from case.json"
        if asset is None:
            for item in images:
                base = posixpath.basename(item["rel"])
                if base and base in figure["caption"]:
                    asset, origin = item["rel"], "image file name mentioned in the caption"
                    break
        if asset is None and index < len(images):
            asset = images[index]["rel"]
            origin = "positional fallback among image inputs (no explicit link found)"
        figure["asset"] = asset
        figure["asset_source"] = origin or "no image asset available in the input"
        if asset and asset in by_rel:
            path = by_rel[asset]["abs"]
            size = common.natural_size(path, by_rel[asset]["kind"])
            if size:
                figure["width"], figure["height"] = size
            else:
                figure["width"] = figure["height"] = None
                figure["size_note"] = "natural size could not be read from the image header"
    return figures


# --------------------------------------------------------------------------------------
# staged inputs (finish runs without --input, so real bytes must survive in <work>)
# --------------------------------------------------------------------------------------

def stage_inputs(input_dir, work_dir, inputs, notes):
    staged = []
    root = os.path.join(work_dir, "inputs")
    for item in inputs:
        destination = os.path.join(root, item["rel"].replace("/", os.sep))
        try:
            sha = common.safe_copy(item["abs"], destination)
        except (OSError, IOError) as exc:
            notes.append("could not stage %s into the work tree: %s" % (item["rel"], exc))
            continue
        if item["sha256"] and sha != item["sha256"]:
            notes.append("staged copy of %s has a different sha256 than the input" % item["rel"])
        staged.append({"rel": item["rel"], "staged": common.relpath_or_abspath(destination, work_dir), "sha256": sha})
    return staged


# --------------------------------------------------------------------------------------
# human-readable review packet
# --------------------------------------------------------------------------------------

def _clip(text, limit):
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[: limit - 3] + "..."


def review_packet(case_id, manifest, template_spec, inventory, catalog_count):
    lines = []
    lines.append("# REVIEW_PACKET — prepare stage (no slides have been produced yet)")
    lines.append("")
    lines.append("- Case: `%s`" % case_id)
    lines.append("- Task type: `%s`" % manifest.get("task_type"))
    lines.append("- Inputs scanned: %d; source units in `slide_catalog.jsonl`: %d" % (len(manifest.get("inputs") or []), catalog_count))
    lines.append("- This stage only *reads* real bytes; it writes no pptx and claims no result.")
    lines.append("")
    lines.append("## 1. File roles (confirmed only from explicit, unique evidence)")
    lines.append("")
    lines.append("| Role | Path(s) |")
    lines.append("|---|---|")
    for role in ("template", "source", "start_deck", "figures"):
        values = (manifest.get("detected") or {}).get(role) or []
        lines.append("| %s | %s |" % (role, ", ".join("`%s`" % v for v in values) if values else "UNVERIFIED"))
    lines.append("")
    lines.append("## 2. Template spec highlights (from real package bytes)")
    lines.append("")
    if template_spec.get("ok"):
        size = template_spec.get("slide_size") or {}
        lines.append("- `%s` sha256 `%s`" % (template_spec.get("template_path"), template_spec.get("sha256")))
        lines.append("- slide size: %sx%s EMU (aspect %s), source `%s`" % (
            size.get("cx"), size.get("cy"), size.get("aspect"), size.get("src")))
        lines.append("- masters: %d; layouts: %d; notes master: %s" % (
            len(template_spec.get("masters") or []), len(template_spec.get("layouts") or []),
            template_spec.get("notes_master") or "none"))
        for master in template_spec.get("masters") or []:
            lines.append("  - `%s` (%s): %d layouts, theme `%s`" % (
                master["part"], master.get("name"), len(master.get("layouts") or []), master.get("theme_rel")))
        for theme in template_spec.get("theme") or []:
            fonts = theme.get("fonts") or {}
            lines.append("  - theme `%s`: major=%s/%s minor=%s/%s accent1=%s" % (
                theme["part"], fonts.get("major_latin"), fonts.get("major_ea"),
                fonts.get("minor_latin"), fonts.get("minor_ea"), (theme.get("colors") or {}).get("accent1")))
        lines.append("")
        lines.append("| layout part | name | type | placeholders (idx:type) |")
        lines.append("|---|---|---|---|")
        for layout in template_spec.get("layouts") or []:
            phs = ", ".join("%s:%s" % (p["idx"], p["ph_type"]) for p in layout.get("placeholders") or [])
            lines.append("| `%s` | %s | %s | %s |" % (layout["part"], layout.get("name"), layout.get("type"), phs))
    else:
        lines.append("- **template is not usable**: %s" % "; ".join(template_spec.get("notes") or ["unknown reason"]))
    lines.append("")
    lines.append("## 3. Source main line")
    lines.append("")
    lines.append("- `%s` (%s), extraction: `%s` / `%s`" % (
        inventory.get("source_path"), inventory.get("kind"),
        (inventory.get("extraction") or {}).get("tool"), (inventory.get("extraction") or {}).get("status")))
    for section in (inventory.get("sections") or [])[:20]:
        lines.append("  - S%d (H%d) %s" % (section["ordinal"], section["level"], _clip(section["title"], 90)))
    if len(inventory.get("sections") or []) > 20:
        lines.append("  - ... %d more sections in `source_inventory.json`" % (len(inventory["sections"]) - 20))
    lines.append("")
    lines.append("### Figures / tables found in the source")
    lines.append("")
    for figure in inventory.get("figures") or []:
        lines.append("- %s → asset `%s` (%sx%s px) `%s`" % (
            figure["label"], figure.get("asset"), figure.get("width"), figure.get("height"), figure.get("source_ref")))
    for table in inventory.get("tables") or []:
        shape = table.get("shape") or {}
        cells = (" [%dx%d real cells from %s]" % (
            shape.get("data_rows") or 0,
            shape.get("cols") or 0,
            table.get("data_source"))
            ) if table.get("data_source") else " [no real cell data]"
        lines.append("- %s → `%s`%s" % (table["label"], _clip(table.get("caption"), 90), cells))
    if not inventory.get("figures") and not inventory.get("tables"):
        lines.append("- none detected in the source text")
    lines.append("")
    lines.append("## 4. Original output contract")
    lines.append("")
    lines.append("- `contract_source`: `%s`; required dotted leaves: %d" % (
        manifest.get("contract_source"), len(manifest.get("required_fields") or [])))
    for path in (manifest.get("required_fields") or [])[:24]:
        lines.append("  - `%s`" % path)
    if len(manifest.get("required_fields") or []) > 24:
        lines.append("  - ... %d more in `case_manifest.json`" % (len(manifest["required_fields"]) - 24))
    lines.append("")
    lines.append("## 5. Conflicts / uncertainty / limitations")
    lines.append("")
    for note in (manifest.get("notes") or [])[:30]:
        lines.append("- %s" % _clip(note, 200))
    if not manifest.get("notes"):
        lines.append("- no conflicts detected while reading the inputs")
    lines.append("")
    lines.append("## 6. Where to look next")
    lines.append("")
    lines.append("- layout/placeholder/theme truth: `template_spec.json`")
    lines.append("- source text units with locators: `source_inventory.json`, `slide_catalog.jsonl`")
    lines.append("- original contract snapshot: `case_manifest.json` (`output_contract`)")
    lines.append("- provenance and staged input copies: `sources.json`, `<work>/inputs/`")
    lines.append("")
    lines.append("> This packet contains no pre-made answers and no fabricated values: every claim above")
    lines.append("> points at a real part, attribute, file byte or source line.")
    lines.append("")
    return "\n".join(lines[:200])


# --------------------------------------------------------------------------------------
# entry points
# --------------------------------------------------------------------------------------

def _toolchain():
    renderer = shutil.which("soffice") or shutil.which("libreoffice")
    pdf_tool, pdf_exe = _detect_pdf_tool()
    chain = [tool for tool, _path in _pdf_tool_chain()]
    return {
        "python": "%d.%d.%d" % sys.version_info[:3],
        "renderer": renderer or "none",
        "pdf_text": pdf_tool or "none",
        "pdf_text_path": pdf_exe or "none",
        "pdf_text_chain": chain,
    }


def discover(input_dir, work_dir, output_dir=None):
    """Prepare stage. Returns a summary dict; writes the discovery artifacts under ``work_dir``."""
    input_dir = os.path.abspath(str(input_dir))
    work_dir = os.path.abspath(str(work_dir))
    output_dir = os.path.abspath(str(output_dir)) if output_dir else None
    common.ensure_dir(work_dir)
    notes = []
    summary = {"ok": False, "status": "needs_human", "work_dir": work_dir,
               "input_dir": input_dir, "notes": notes, "artifacts": {}}

    if not os.path.isdir(input_dir):
        notes.append("input directory does not exist: %s" % input_dir)
        return summary

    inputs = common.scan_inputs(input_dir)
    by_rel = {item["rel"]: item for item in inputs}
    case_path = os.path.join(input_dir, "case.json")
    task_path = os.path.join(input_dir, "task.md")

    case = {}
    case_id = None
    contract = {}
    contract_source = "absent"
    instruction_text = ""
    constraints = {}
    budget = dict(DEFAULT_BUDGET)
    case_error = None
    if not os.path.isfile(case_path):
        case_error = "missing required input file case.json"
        notes.append(case_error)
    else:
        try:
            with open(case_path, "r", encoding="utf-8-sig") as handle:
                case = json.load(handle)
        except (OSError, ValueError) as exc:
            case_error = "case.json could not be parsed: %s: %s" % (type(exc).__name__, exc)
            notes.append(case_error)
            case = {}
        if not isinstance(case, dict):
            case_error = "case.json root is not a JSON object"
            notes.append(case_error)
            case = {}

    task_text = ""
    if os.path.isfile(task_path):
        try:
            with open(task_path, "r", encoding="utf-8-sig", errors="replace") as handle:
                task_text = handle.read()
        except OSError as exc:
            notes.append("task.md could not be read: %s" % exc)

    if case:
        case_id = _first_str(case, ("case_id", "caseId", "id", "task_id", "taskId"))
        if case_id is None:
            for wrapper in ("identification", "case", "task", "meta"):
                nested = case.get(wrapper)
                if isinstance(nested, dict):
                    case_id = _first_str(nested, ("case_id", "caseId", "id"))
                    if case_id:
                        break
        if case_id is None:
            notes.append("case.json carries no case_id under any known key; identification.case_id will be null.")
        instruction_text = _first_str(case, ("instructions", "instruction", "requirement", "task", "prompt")) or ""
        found, where = _find_output_contract(case)
        if found:
            contract, contract_source = found, "case.json"
        budget_value = case.get("budget") if isinstance(case.get("budget"), dict) else None
        if budget_value:
            budget = {
                "model_requests": budget_value.get("model_requests", DEFAULT_BUDGET["model_requests"]),
                "seconds": budget_value.get("seconds", DEFAULT_BUDGET["seconds"]),
            }
        constraints = case.get("constraints") if isinstance(case.get("constraints"), dict) else {}
    if not contract and task_text:
        fallback = _contract_from_task_md(task_text)
        if fallback is not None:
            contract, contract_source = fallback.get("output_contract") or {}, "task.md"
            notes.append("output_contract was read from a fenced JSON block in task.md, not case.json.")
    if not contract:
        notes.append("no output_contract could be located; required_fields stays empty (nothing is invented).")
    if not instruction_text and task_text:
        instruction_text = task_text.strip()
        notes.append("instructions were taken from task.md because case.json carried none.")
    if case.get("constraints"):
        for key in ("language", "aspect", "target_slide_count", "caption_numbering"):
            if key in case["constraints"]:
                constraints.setdefault(key, case["constraints"][key])

    required_fields = _dotted_leaves(contract) if contract else []

    detected, by_rel = classify_roles(case if case else {}, inputs, notes)
    if case.get("constraints"):
        for key in ("masters_min", "layouts_min"):
            if key in case["constraints"]:
                constraints.setdefault(key, case["constraints"][key])

    # ---- template spec -----------------------------------------------------------------
    template_rel = detected["template"][0] if len(detected["template"]) == 1 else None
    if template_rel and template_rel in by_rel:
        item = by_rel[template_rel]
        template_spec = parse_template(item["abs"], template_rel, item["sha256"])
    else:
        template_spec = {
            "template_path": template_rel, "sha256": None, "ok": False,
            "slide_size": None, "masters": [], "layouts": [], "theme": [],
            "notes_master": None, "evidence": [],
            "notes": ["no unique template candidate; template spec was not parsed"],
        }
        notes.append("template spec stays empty because no unique template candidate was confirmed.")

    # ---- start deck spec (optional) ----------------------------------------------------
    start_rel = detected["start_deck"][0] if len(detected["start_deck"]) == 1 else None
    start_spec = None
    if start_rel and start_rel in by_rel:
        item = by_rel[start_rel]
        start_spec = start_deck_spec(item["abs"], start_rel, item["sha256"])

    # ---- source inventory --------------------------------------------------------------
    source_rel = detected["source"][0] if len(detected["source"]) == 1 else None
    sections, units, figures, tables = [], [], [], []
    extraction = {"tool": "none", "status": "failed", "detail": "no unique source was confirmed"}
    if source_rel and source_rel in by_rel:
        item = by_rel[source_rel]
        sections, units, figures, tables, extraction = parse_source(item["abs"], source_rel, item["kind"])
        if extraction.get("status") != "ok":
            notes.append("source text extraction failed honestly: %s" % extraction.get("detail"))
    else:
        notes.append("no unique source candidate; source_inventory carries no units.")
    figures = _attach_figure_assets(figures, inputs, detected)
    if figures and all(figure.get("asset") is None for figure in figures):
        notes.append("figure captions were found but no image asset could be attached.")

    inventory = {
        "source_path": source_rel,
        "kind": by_rel[source_rel]["kind"] if source_rel in by_rel else None,
        "sha256": by_rel[source_rel]["sha256"] if source_rel in by_rel else None,
        "sections": sections,
        "units": units,
        "figures": figures,
        "tables": tables,
        "text_available": bool(units) and extraction.get("status") == "ok",
        "extraction": extraction,
        "notes": [n for n in notes if "source" in n or "figure" in n or "extract" in n],
    }

    staged = stage_inputs(input_dir, work_dir, inputs, notes)
    if staged:
        notes.append("all readable inputs were staged byte-identical under <work>/inputs/ so that "
                     "`finish` can rebuild the deck without reading a live /input again.")

    manifest = {
        "case_id": case_id,
        "task_type": "academic_ppt_template_adaptation",
        "inputs": [{k: item.get(k) for k in ("rel", "abs", "kind", "size", "sha256")} for item in inputs],
        "output_contract": copy.deepcopy(contract),
        "required_fields": required_fields,
        "contract_source": contract_source,
        "instruction_text": instruction_text,
        "budget": budget,
        "detected": detected,
        "constraints": constraints,
        "staged_inputs": staged,
        "notes": notes,
    }

    artifacts = {}
    artifacts["case_manifest"] = os.path.join(work_dir, "case_manifest.json")
    common.write_json(artifacts["case_manifest"], manifest)
    artifacts["template_spec"] = os.path.join(work_dir, "template_spec.json")
    common.write_json(artifacts["template_spec"], template_spec)
    artifacts["source_inventory"] = os.path.join(work_dir, "source_inventory.json")
    common.write_json(artifacts["source_inventory"], inventory)
    artifacts["slide_catalog"] = os.path.join(work_dir, "slide_catalog.jsonl")
    if os.path.isfile(artifacts["slide_catalog"]):
        os.remove(artifacts["slide_catalog"])
    for record in units:
        common.append_jsonl(artifacts["slide_catalog"], record)
    artifacts["review_packet"] = os.path.join(work_dir, "REVIEW_PACKET.md")
    common.write_text(artifacts["review_packet"], review_packet(case_id, manifest, template_spec, inventory, len(units)))
    if start_spec is not None:
        artifacts["start_deck_spec"] = os.path.join(work_dir, "start_deck_spec.json")
        common.write_json(artifacts["start_deck_spec"], start_spec)

    sources = {
        "inputs": [{"rel": item["rel"], "kind": item["kind"], "size": item.get("size"), "sha256": item.get("sha256")} for item in inputs],
        "template_sha256": template_spec.get("sha256"),
        "source_sha256": inventory.get("sha256"),
        "deck_sha256": None,
        "skill": {
            "name": common.SKILL_NAME,
            "version": common.SKILL_VERSION,
            "scripts_sha256": {"common.py": common.sha256_file(os.path.join(_HERE, "common.py"))
                               if os.path.isfile(os.path.join(_HERE, "common.py")) else None},
        },
        "toolchain": _toolchain(),
        "reproduce": {
            "prepare": "python -B scripts/run.py prepare --input /input --work /app/work --output /app/output",
            "finish": "python -B scripts/run.py finish --work /app/work --output /app/output",
            "standalone_discover": "python -B scripts/discover.py --input /input --work /app/work",
        },
        "staged_inputs": staged,
        "uncertainty": [n for n in notes if "unconfirmed" in n or "UNVERIFIED" in n or "case_id" in n],
        "limitations": [
            "prepare inspects inputs only; no slide is produced and no fidelity/format/render claim is made.",
            "roles, theme semantics and caption-to-asset matches are read from real evidence and may stay unconfirmed.",
        ],
    }
    artifacts["sources"] = os.path.join(work_dir, "sources.json")
    common.write_json(artifacts["sources"], sources)

    ok = case_error is None and template_spec.get("ok") is True and extraction.get("status") == "ok" and bool(units)
    summary.update({
        "ok": ok,
        "status": "ok" if ok else "needs_human",
        "case_id": case_id,
        "case_error": case_error or None,
        "input_count": len(inputs),
        "template_ok": bool(template_spec.get("ok")),
        "masters": len(template_spec.get("masters") or []),
        "layouts": len(template_spec.get("layouts") or []),
        "source_units": len(units),
        "figures": len(figures),
        "tables": len(tables),
        "detected": detected,
        "extraction": extraction,
        "notes": notes,
        "artifacts": artifacts,
    })
    if not ok:
        summary["reason"] = case_error or "discovery did not confirm a usable template and source"
    return summary


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", required=True, help="input root (case.json + task.md + materials)")
    parser.add_argument("--work", required=True, help="writable directory for discovery artifacts")
    parser.add_argument("--output", default=None, help="writable output dir (only recorded, not written here)")
    args = parser.parse_args(argv)
    summary = discover(args.input, args.work, args.output)
    sys.stdout.write(json.dumps(summary, ensure_ascii=False, indent=2, default=str) + "\n")
    if not os.path.isdir(args.input):
        return 2
    # CONTRACT §4: an absent/unreadable case.json is an argument-class failure (exit 2).
    # A *readable* case that merely degrades (broken template, failed extraction) still
    # exits 0: the degradation is visible in the artifacts and summary["ok"] is False.
    return 2 if summary.get("case_error") else 0


if __name__ == "__main__":
    raise SystemExit(main())
