#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""adapt_pptx.py — build an editable deck that *really* inherits the organizer template.

Owner: lead (work/CONTRACT.md §1).  Pure standard library; offline; POSIX/Windows safe.

What "really inherits" means here (and how it is checkable):
  * every slide/document/master/layout/theme part of the template is copied **byte for byte**;
  * each produced slide has a relationship to a real ``ppt/slideLayouts/slideLayoutN.xml``;
  * each produced placeholder carries the ``<p:ph type=.. idx=..>`` pair that the chosen
    layout actually declares, and its ``<a:xfrm>`` is copied from that layout (or from the
    master when the layout inherits geometry) so overflow can be recomputed later;
  * text runs never set ``latin``/``ea`` typefaces, so theme fonts stay in force;
  * images are scaled with a single factor (``min`` contain), never stretched;
  * a speaker-notes master is synthesised when the template has none, so notes are real.

Reads ``<work>/{case_manifest,template_spec,source_inventory,slides_plan}.json`` and writes
``<output>/<deck_name>`` plus ``<work>/{adaptation_actions.json,change_records.jsonl}``.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import zipfile
from xml.etree import ElementTree as ET

import common as C

CT = {
    "sld": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "notes": "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    "notesMaster": "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
}
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

IMG_EXT = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
           "bmp": "image/bmp", "tiff": "image/tiff", "webp": "image/webp"}

# name -> ladder of sizes in 1/100 pt (largest first)
LADDER = {
    "title": [3200, 2800, 2400],
    "body": [2000, 1800, 1600, 1400],
    "caption": [1400, 1200],
    "table": [1400, 1200],
}
DEFAULT_SIZE = {"title": 3200, "body": 1800, "caption": 1200, "table": 1200}

TEXT_INSET_L = 91440
TEXT_INSET_T = 45720

# The adapter must be *at least as strict* as the independent auditor, otherwise a deck can
# pass here and still be reported as overflowing there.  Width gets a safety margin and the
# capacity is computed with a slightly wider line than the auditor's 1.2.
SAFE_W = 0.97
SAFE_LINE_SPACING = 1.25

_MD = [
    (re.compile(r"\*\*(.+?)\*\*", re.S), r"\1"),
    (re.compile(r"__(.+?)__", re.S), r"\1"),
    (re.compile(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", re.S), r"\1"),
    (re.compile(r"`([^`]+)`"), r"\1"),
    (re.compile(r"^\s{0,3}>\s?", re.M), ""),
    (re.compile(r"^\s{0,3}[-*+]\s+", re.M), ""),
    (re.compile(r"^\s{0,3}#{1,6}\s+", re.M), ""),
    (re.compile(r"\\n"), " "),
]


def plain_text(s):
    """Strip markdown emphasis/list/quote markers; keeps every word and number intact."""
    out = s or ""
    for pat, rep in _MD:
        out = pat.sub(rep, out)
    return re.sub(r"\s+", " ", out).strip()


def _register():
    ET.register_namespace("p", C.NS["p"])
    ET.register_namespace("a", C.NS["a"])
    ET.register_namespace("r", C.NS["r"])
    ET.register_namespace("", PKG_REL)


_register()


def q(tag):
    prefix, _, local = tag.partition(":")
    return "{%s}%s" % (C.NS[prefix], local)


def sub(parent, tag, **attrs):
    e = ET.SubElement(parent, q(tag))
    for k, v in attrs.items():
        if v is not None:
            e.set(k, str(v))
    return e


# ======================================================================================
# 1. template reading
# ======================================================================================

class Template:
    def __init__(self, path):
        self.path = path
        self.sha256 = C.sha256_file(path)
        self.parts = {}
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                self.parts[info.filename] = z.read(info.filename)
        self.masters = sorted(n for n in self.parts if re.match(r"^ppt/slideMasters/slideMaster\d+\.xml$", n))
        self.layouts = sorted(
            (n for n in self.parts if re.match(r"^ppt/slideLayouts/slideLayout\d+\.xml$", n)),
            key=lambda n: int(re.search(r"slideLayout(\d+)", n).group(1)))
        self.themes = sorted(n for n in self.parts if re.match(r"^ppt/theme/theme\d+\.xml$", n))
        self.notes_masters = sorted(n for n in self.parts if re.match(r"^ppt/notesMasters/notesMaster\d+\.xml$", n))
        self.presentation = self.parts.get("ppt/presentation.xml")
        self.pres_rels = self.parts.get("ppt/_rels/presentation.xml.rels")
        self.content_types = self.parts.get("[Content_Types].xml")
        self._layout_cache = {}

    def rels_of(self, part):
        base = os.path.dirname(part)
        name = os.path.basename(part)
        rel = "%s/_rels/%s.rels" % (base, name)
        raw = self.parts.get(rel)
        if not raw:
            return {}
        root = ET.fromstring(raw)
        out = {}
        for r in root.findall("{%s}Relationship" % PKG_REL):
            out[r.get("Id")] = {"type": r.get("Type"), "target": r.get("Target"),
                                "mode": r.get("TargetMode")}
        return out

    def resolve(self, part, target):
        if target.startswith("/"):
            return target.lstrip("/")
        return os.path.normpath(os.path.join(os.path.dirname(part), target)).replace("\\", "/")

    def layout_meta(self, layout_part):
        """Placeholders (by idx) with geometry + declared text size, plus the layout name."""
        if layout_part in self._layout_cache:
            return self._layout_cache[layout_part]
        root = ET.fromstring(self.parts[layout_part])
        meta = {"part": layout_part, "name": (root.get("name") or "").strip(),
                "type": root.get("type") or "", "placeholders": {}, "shapes": []}
        master_ph = {}
        for r in self.rels_of(layout_part).values():
            if r["type"].endswith("/slideMaster"):
                mp = self.resolve(layout_part, r["target"])
                master_ph = self._master_placeholders(mp)
                meta["master"] = mp
                break
        tree = root.find("p:cSld/p:spTree", C.NS)
        if tree is not None:
            for sp in list(tree):
                if C.localname(sp.tag) not in ("sp", "graphicFrame", "pic"):
                    continue
                ph = self._find_ph(sp)
                if ph is None:
                    continue
                key = self._ph_key(ph)
                info = self._shape_geometry(sp)
                info["name"] = self._shape_name(sp)
                info["ph_type"] = ph.get("type") or "body"
                info["ph_idx"] = ph.get("idx")
                info["size"] = self._declared_text_size(sp)
                info["alt_rect"], info["alt_size"] = info.get("xfrm"), None
                if not info.get("xfrm") and key in master_ph:
                    info["xfrm"] = master_ph[key].get("xfrm")
                    info["xfrm_source"] = "master"
                    info["size"] = info.get("size") or master_ph[key].get("size")
                else:
                    info["xfrm_source"] = "layout" if info.get("xfrm") else None
                info["inherits_from_master"] = key in master_ph
                meta["placeholders"][key] = info
        self._layout_cache[layout_part] = meta
        return meta

    def _master_placeholders(self, master_part):
        out = {}
        raw = self.parts.get(master_part)
        if not raw:
            return out
        root = ET.fromstring(raw)
        tree = root.find("p:cSld/p:spTree", C.NS)
        if tree is None:
            return out
        for sp in list(tree):
            ph = self._find_ph(sp)
            if ph is None:
                continue
            info = self._shape_geometry(sp)
            info["size"] = self._declared_text_size(sp)
            out[self._ph_key(ph)] = info
        return out

    @staticmethod
    def _find_ph(sp):
        nv = sp.find("p:nvSpPr/p:nvPr", C.NS)
        if nv is None:
            nv = sp.find("p:nvGraphicFramePr/p:nvPr", C.NS)
        if nv is None:
            nv = sp.find("p:nvPicPr/p:nvPr", C.NS)
        return nv.find("p:ph", C.NS) if nv is not None else None

    @staticmethod
    def _ph_key(ph):
        t = ph.get("type") or "body"
        idx = ph.get("idx")
        if idx is None:
            idx = "0" if t in ("title", "ctrTitle") else None
        return (t, idx)

    @staticmethod
    def _shape_name(sp):
        nv = sp.find("p:nvSpPr/p:cNvPr", C.NS)
        if nv is None:
            nv = sp.find("p:nvPicPr/p:cNvPr", C.NS)
        if nv is None:
            nv = sp.find("p:nvGraphicFramePr/p:cNvPr", C.NS)
        return nv.get("name") if nv is not None else None

    @staticmethod
    def _shape_geometry(sp):
        xfrm = sp.find("p:spPr/a:xfrm", C.NS)
        if xfrm is None:
            xfrm = sp.find("p:grpSpPr/a:xfrm", C.NS)
        out = {"xfrm": None, "body_pr": None}
        if xfrm is not None:
            off = xfrm.find("a:off", C.NS)
            ext = xfrm.find("a:ext", C.NS)
            if off is not None and ext is not None:
                out["xfrm"] = {"x": int(off.get("x", 0)), "y": int(off.get("y", 0)),
                               "cx": int(ext.get("cx", 0)), "cy": int(ext.get("cy", 0))}
        body_pr = sp.find("p:txBody/a:bodyPr", C.NS)
        if body_pr is not None:
            out["body_pr"] = {k: body_pr.get(k) for k in
                              ("anchor", "lIns", "tIns", "rIns", "bIns", "wrap", "rot") if body_pr.get(k) is not None}
            for child in body_pr:
                out["body_pr"]["autofit"] = C.localname(child.tag)
        return out

    @staticmethod
    def _declared_text_size(sp):
        d = sp.find("p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr", C.NS)
        if d is not None and d.get("sz"):
            try:
                return int(d.get("sz"))
            except ValueError:
                return None
        return None

    def slide_size(self):
        if not self.presentation:
            return None
        root = ET.fromstring(self.presentation)
        sz = root.find("p:sldSz", C.NS)
        if sz is None:
            return None
        return {"cx": int(sz.get("cx")), "cy": int(sz.get("cy")), "type": sz.get("type")}

    def theme_fonts_colors(self):
        """For audit: what the template theme actually declares."""
        out = []
        for t in self.themes:
            root = ET.fromstring(self.parts[t])
            fs = root.find(".//a:fontScheme", C.NS)
            cs = root.find(".//a:clrScheme", C.NS)
            rec = {"part": t, "fonts": {}, "colors": {}}
            if fs is not None:
                for kind in ("majorFont", "minorFont"):
                    node = fs.find("a:" + kind, C.NS)
                    if node is None:
                        continue
                    for script in ("latin", "ea", "cs"):
                        el = node.find("a:" + script, C.NS)
                        if el is not None:
                            rec["fonts"]["%s_%s" % ("major" if kind == "majorFont" else "minor", script)] = el.get("typeface")
            if cs is not None:
                for child in cs:
                    name = C.localname(child.tag)
                    srgb = child.find("a:srgbClr", C.NS)
                    sysc = child.find("a:sysClr", C.NS)
                    if srgb is not None:
                        rec["colors"][name] = srgb.get("val")
                    elif sysc is not None:
                        rec["colors"][name] = sysc.get("lastClr") or sysc.get("val")
            out.append(rec)
        return out


# ======================================================================================
# 2. slide XML construction
# ======================================================================================

def _sp_tree():
    tree = ET.Element(q("p:spTree"))
    nv = sub(tree, "p:nvGrpSpPr")
    sub(nv, "p:cNvPr", id="1", name="")
    sub(nv, "p:cNvGrpSpPr")
    sub(nv, "p:nvPr")
    g = sub(tree, "p:grpSpPr")
    xfrm = sub(g, "a:xfrm")
    sub(xfrm, "a:off", x=0, y=0)
    sub(xfrm, "a:ext", cx=0, cy=0)
    sub(xfrm, "a:chOff", x=0, y=0)
    sub(xfrm, "a:chExt", cx=0, cy=0)
    return tree


def _add_text_paragraph(txbody, text, lvl, sz, bold=False, lang="zh-CN"):
    p = sub(txbody, "a:p")
    attrs = {"algn": "l"}
    if lvl:
        attrs["lvl"] = lvl
    sub(p, "a:pPr", **attrs)
    if text == "":
        sub(p, "a:endParaRPr", lang=lang)
        return p
    r = sub(p, "a:r")
    r_attrs = {"lang": lang, "dirty": "0"}
    if sz:
        r_attrs["sz"] = sz
    if bold:
        r_attrs["b"] = "1"
    sub(r, "a:rPr", **r_attrs)
    sub(r, "a:t").text = text
    return p


def _placeholder_sp(shape_id, name, ph_type, ph_idx, xfrm, paragraphs, size, *, bold=False,
                    body_pr=None, substituted=False):
    sp = ET.Element(q("p:sp"))
    nv = sub(sp, "p:nvSpPr")
    sub(nv, "p:cNvPr", id=shape_id, name=name)
    cnv = sub(nv, "p:cNvSpPr")
    if not substituted:
        sub(cnv, "a:spLocks", noGrp="1")
    nvpr = sub(nv, "p:nvPr")
    if not substituted:
        attrs = {}
        if ph_type and ph_type != "body":
            attrs["type"] = ph_type
        if ph_idx is not None:
            attrs["idx"] = ph_idx
        sub(nvpr, "p:ph", **attrs)
    sp_xfrm = sub(sp, "p:spPr")
    if xfrm:
        x = sub(sp_xfrm, "a:xfrm")
        sub(x, "a:off", x=xfrm["x"], y=xfrm["y"])
        sub(x, "a:ext", cx=xfrm["cx"], cy=xfrm["cy"])
    sub(sp_xfrm, "a:prstGeom", prst="rect")
    tx = sub(sp, "p:txBody")
    bp = sub(tx, "a:bodyPr", wrap="square", lIns=TEXT_INSET_L, tIns=TEXT_INSET_T,
             rIns=TEXT_INSET_L, bIns=TEXT_INSET_T)
    if body_pr and body_pr.get("anchor"):
        bp.set("anchor", body_pr["anchor"])
    sub(bp, "a:normAutofit")
    sub(tx, "a:lstStyle")
    for para in paragraphs:
        _add_text_paragraph(tx, plain_text(para.get("text", "")), para.get("level", 0), size, bold=bold)
    if not paragraphs:
        _add_text_paragraph(tx, "", 0, size, bold=bold)
    return sp


def _picture_sp(shape_id, name, rid, xfrm, descr=None):
    pic = ET.Element(q("p:pic"))
    nv = sub(pic, "p:nvPicPr")
    sub(nv, "p:cNvPr", id=shape_id, name=name, descr=descr)
    cnv = sub(nv, "p:cNvPicPr")
    sub(cnv, "a:picLocks", noChangeAspect="1")
    sub(nv, "p:nvPr")
    fill = sub(pic, "p:blipFill")
    sub(fill, "a:blip", **{q("r:embed"): rid})
    stretch = sub(fill, "a:stretch")
    sub(stretch, "a:fillRect")
    sp = sub(pic, "p:spPr")
    x = sub(sp, "a:xfrm")
    sub(x, "a:off", x=xfrm["x"], y=xfrm["y"])
    sub(x, "a:ext", cx=xfrm["cx"], cy=xfrm["cy"])
    sub(sp, "a:prstGeom", prst="rect")
    return pic


def slide_xml(sp_tree):
    sld = ET.Element(q("p:sld"))
    cSld = sub(sld, "p:cSld")
    cSld.append(sp_tree)
    ovr = sub(sld, "p:clrMapOvr")
    sub(ovr, "a:masterClrMapping")
    return C.xml_bytes(sld)


def notes_xml(paragraphs):
    notes = ET.Element(q("p:notes"))
    cSld = sub(notes, "p:cSld")
    tree = _sp_tree()
    # notes body placeholder (idx 1 per the notes master convention)
    sp = _placeholder_sp(2, "Notes Placeholder 1", "body", "1", None,
                         [{"text": t, "level": 0} for t in paragraphs], None)
    tree.append(sp)
    # slide image placeholder so the notes page keeps its layout
    sp2 = _placeholder_sp(3, "Slide Image Placeholder 1", "sldImg", None, None, [], None)
    tree.append(sp2)
    cSld.append(tree)
    ovr = sub(notes, "p:clrMapOvr")
    sub(ovr, "a:masterClrMapping")
    return C.xml_bytes(notes)


NOTES_MASTER = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1714500"/><a:ext cx="5486400" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="2743200" y="457200"/><a:ext cx="2743200" cy="2057400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" indent="0" algn="l"><a:defRPr sz="1200"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>"""


# ======================================================================================
# 3. layout choice + text fitting
# ======================================================================================

def choose_layout(tpl, slide, default_part=None):
    want_part = slide.get("layout_part")
    if want_part and want_part in tpl.parts:
        return want_part
    want_name = C.norm_key(slide.get("layout_name") or "")
    if want_name:
        for lp in tpl.layouts:
            meta = tpl.layout_meta(lp)
            if C.norm_key(meta["name"]) == want_name:
                return lp
    role = (slide.get("role") or "").lower()
    if role in ("title", "cover") or slide.get("n") == 1:
        for lp in tpl.layouts:
            if (tpl.layout_meta(lp)["type"] or "").lower() in ("title", "titleonly") or \
               any(k[0] == "ctrTitle" for k in tpl.layout_meta(lp)["placeholders"]):
                return lp
    for lp in tpl.layouts:
        if any(k[0] in ("body", "obj") for k in tpl.layout_meta(lp)["placeholders"]):
            return lp
    return tpl.layouts[0] if tpl.layouts else None


def pick_placeholder(meta, kinds):
    """Return (key, info) preferring the declared order of ``kinds``."""
    for kind in kinds:
        for key, info in meta["placeholders"].items():
            if key[0] == kind:
                return key, info
    return None, None


def _usable_rect(info):
    if not info or not info.get("xfrm"):
        return None
    x = info["xfrm"]
    lp = float(info.get("body_pr", {}).get("lIns", TEXT_INSET_L)) if info.get("body_pr") else TEXT_INSET_L
    rp = float(info.get("body_pr", {}).get("rIns", TEXT_INSET_L)) if info.get("body_pr") else TEXT_INSET_L
    tp = float(info.get("body_pr", {}).get("tIns", TEXT_INSET_T)) if info.get("body_pr") else TEXT_INSET_T
    bp = float(info.get("body_pr", {}).get("bIns", TEXT_INSET_T)) if info.get("body_pr") else TEXT_INSET_T
    return {"x": x["x"], "y": x["y"], "cx": x["cx"], "cy": x["cy"],
            "w_pt": C.emu_to_pt(x["cx"] - lp - rp), "h_pt": C.emu_to_pt(x["cy"] - tp - bp)}


def fit_size(paras, box, ladder, floor):
    """Walk the ladder largest-first; return (size, metrics, history) or (None, ..) if overflow."""
    history = []
    if box is None or box["w_pt"] <= 0 or box["h_pt"] <= 0:
        return None, {"fits": None, "detail": "no usable geometry"}, history
    for size in ladder:
        total_lines = 0
        max_w = 0.0
        for p in paras:
            ind = 12.0 * float(p.get("level", 0) or 0)
            m = C.estimate_wrapped_lines(plain_text(p.get("text", "")), size / 100.0,
                                         max(1.0, (box["w_pt"] - ind) * SAFE_W),
                                         line_spacing=SAFE_LINE_SPACING)
            total_lines += max(1, m["lines"])
            max_w = max(max_w, m["max_line_w_pt"])
        cap = C.capacity_lines(box["h_pt"], size / 100.0, SAFE_LINE_SPACING)
        rec = {"size": size, "lines": total_lines, "capacity_lines": cap,
               "width_ok": max_w <= box["w_pt"] + 0.5}
        history.append(rec)
        if total_lines <= cap and rec["width_ok"]:
            return size, {"fits": True, "lines": total_lines, "capacity_lines": cap,
                          "max_line_w_pt": max_w, "box_w_pt": box["w_pt"], "box_h_pt": box["h_pt"]}, history
        if size <= floor:
            break
    return None, {"fits": False, "lines": history[-1]["lines"] if history else None,
                  "capacity_lines": history[-1]["capacity_lines"] if history else None,
                  "box_w_pt": box["w_pt"], "box_h_pt": box["h_pt"]}, history


def pack_paragraphs(paras, box, size, floor_slack=1.0):
    """Split a bullet list into chunks that each fit the box at the given size."""
    if box is None or box["w_pt"] <= 0 or box["h_pt"] <= 0:
        return [paras]
    cap = C.capacity_lines(box["h_pt"], size / 100.0, SAFE_LINE_SPACING) * floor_slack
    chunks, cur, used = [], [], 0
    for p in paras:
        ind = 12.0 * float(p.get("level", 0) or 0)
        m = C.estimate_wrapped_lines(plain_text(p.get("text", "")), size / 100.0,
                                     max(1.0, (box["w_pt"] - ind) * SAFE_W),
                                     line_spacing=SAFE_LINE_SPACING)
        need = max(1, m["lines"])
        if cur and used + need > cap:
            chunks.append(cur)
            cur, used = [], 0
        cur.append(p)
        used += need
    if cur:
        chunks.append(cur)
    return chunks or [paras]


def contain_rect(rect, iw, ih):
    """Largest centred box with the image's aspect ratio inside ``rect`` (never stretches)."""
    if not iw or not ih or rect["cx"] <= 0 or rect["cy"] <= 0:
        return {"x": rect["x"], "y": rect["y"], "cx": rect["cx"], "cy": rect["cy"],
                "scale": None, "aspect_delta": None, "letterbox": True}
    scale = min(rect["cx"] / float(iw), rect["cy"] / float(ih))
    cx, cy = int(iw * scale), int(ih * scale)
    x = rect["x"] + (rect["cx"] - cx) // 2
    y = rect["y"] + (rect["cy"] - cy) // 2
    return {"x": x, "y": y, "cx": cx, "cy": cy, "scale": round(scale, 6),
            "aspect_delta": 0.0, "letterbox": (cx < rect["cx"] or cy < rect["cy"])}


# ======================================================================================
# 4. the adapter
# ======================================================================================

class Adapter:
    def __init__(self, work_dir, output_dir, tpl_path, spec, inventory, plan, deck_name):
        self.work = work_dir
        self.output = output_dir
        self.tpl = Template(tpl_path)
        self.spec = spec or {}
        self.inventory = inventory or {}
        self.plan = plan or {}
        self.deck_name = deck_name
        self.actions = []
        self.records = []
        self.slides = []          # (part, layout_part, rels, meta)
        self.media = {}           # source path -> (part, sha)
        self.media_seq = 0
        self.shape_seq = 1        # id 1 is the group shape in _sp_tree()
        self.fig_index = {}
        for f in (self.inventory.get("figures") or []):
            self.fig_index[f.get("id")] = f

    # ---- bookkeeping -------------------------------------------------------------
    def action(self, kind, slide, target, before, after, evidence, note=""):
        aid = "a%d" % (len(self.actions) + 1)
        self.actions.append({"id": aid, "kind": kind, "slide": slide, "target": target,
                             "before": before, "after": after, "evidence": evidence, "note": note})
        return aid

    def record(self, slide, kind, target, before, after, evidence):
        self.records.append({"ts": C.now_iso(), "slide": slide, "kind": kind, "target": target,
                             "before": before, "after": after, "evidence": evidence})

    def _next_shape_id(self):
        self.shape_seq += 1
        return self.shape_seq

    def add_media(self, path):
        if path in self.media:
            return self.media[path]
        kind = C.guess_kind(path)
        ext = "png" if kind == "png" else ("jpeg" if kind in ("jpg", "jpeg") else
                                           (kind if kind in IMG_EXT else "png"))
        self.media_seq += 1
        part = "ppt/media/adaptImage%d.%s" % (self.media_seq, ext)
        data = open(path, "rb").read()
        self.media[path] = {"part": part, "sha256": C.sha256_bytes(data), "bytes": len(data),
                            "ext": ext, "kind": kind}
        return self.media[path]

    # ---- layout geometry ---------------------------------------------------------
    def body_rect(self, layout_part):
        meta = self.tpl.layout_meta(layout_part)
        for kinds in (("body", "obj", "subTitle"), ("title", "ctrTitle")):
            key, info = pick_placeholder(meta, kinds)
            if info:
                r = _usable_rect(info)
                if r:
                    return r, key, info
        return None, None, None

    # ---- one slide ---------------------------------------------------------------
    def build_slide(self, plan_slide, *, continuation_of=None, paragraphs=None):
        tpl = self.tpl
        layout_part = choose_layout(tpl, plan_slide)
        meta = tpl.layout_meta(layout_part)
        size = tpl.slide_size() or {"cx": 12192000, "cy": 6858000}
        floor = (self.plan.get("font_ladder") or {}).get("floor") or {}
        n = plan_slide.get("n")
        title_key, title_info = pick_placeholder(meta, ("ctrTitle", "title"))
        body_key, body_info = pick_placeholder(meta, ("body", "obj", "subTitle"))
        body_rect, _, _ = self.body_rect(layout_part)

        title_text = plan_slide.get("title") or ""
        if continuation_of is not None:
            title_text = "%s（续）" % continuation_of
        bullets = paragraphs if paragraphs is not None else (plan_slide.get("bullets") or [])
        figures = plan_slide.get("figures") or []

        # --- title pass: may extend the title box and push the body down ----------
        # Titles in real papers are long sentences.  The task explicitly names "title too long"
        # as a case to handle with an explainable ladder, so when the layout's own title box
        # cannot hold the title at the readable floor we extend the box (keeping the text and
        # the floor size) and give the body the remaining area, recording both changes.
        tl_ladder = ((self.plan.get("font_ladder") or {}).get("title") or LADDER["title"])
        tbase = (title_info or {}).get("size")
        if tbase:
            tl_ladder = [s for s in tl_ladder if s <= tbase] or [tbase]
        floor_title = int(floor.get("title") or C.FONT_FLOOR["title"])
        tr = _usable_rect(title_info) if title_info else None
        title_xfrm = dict((title_info or {}).get("xfrm") or {}) or None
        title_size, title_metrics, title_hist = fit_size(
            [{"text": title_text, "level": 0}], tr, tl_ladder, floor_title)
        body_top_shift = 0
        if title_size is None:
            title_size = min(floor_title, tl_ladder[-1])
            title_metrics = dict(title_metrics or {})
            title_metrics["action"] = "title_floor"
            need_h = int(int(title_metrics.get("lines") or 0) * (title_size / 100.0)
                         * SAFE_LINE_SPACING * C.EMU_PER_PT) + 2 * TEXT_INSET_T
            if title_xfrm and need_h > title_xfrm["cy"]:
                cap_h = int(size["cy"] * 0.4)
                new_cy = min(need_h, cap_h)
                before_cy = title_xfrm["cy"]
                title_xfrm["cy"] = new_cy
                body_top_shift = new_cy - before_cy
                self.action("title_box_extended", n, "ppt/slides/slide%d.xml" % n,
                            "title box cy=%d at floor size %d (measured %s lines)"
                            % (before_cy, title_size, title_metrics.get("lines")),
                            "title box cy=%d (+%d EMU); body area pushed down by the same amount"
                            % (new_cy, body_top_shift),
                            ["common.estimate_wrapped_lines", "template_spec.layouts[%s]" % layout_part],
                            "the layout's title box cannot hold this title at the readable floor; "
                            "extending the box keeps the text intact instead of shrinking it below the floor")
                self.record(n, "title_box_extended", "ppt/slides/slide%d.xml" % n,
                            {"cy": before_cy}, {"cy": new_cy, "sz": title_size}, "measured overflow")
        elif title_hist and len(title_hist) > 1:
            title_metrics["downgraded_from"] = title_hist[0]["size"]
            self.action("font_ladder", n, "ppt/slides/slide%d.xml" % n,
                        "declared title size %d" % title_hist[0]["size"],
                        "title size %d after %d ladder step(s)" % (title_size, len(title_hist) - 1),
                        [json.dumps(title_hist)], "font ladder applied before any geometry change")

        if body_top_shift and body_rect:
            bottom = body_rect["y"] + body_rect["cy"]
            new_y = min(body_rect["y"] + body_top_shift, bottom - 200000)
            body_rect = dict(body_rect)
            body_rect["y"] = new_y
            body_rect["cy"] = max(1, bottom - new_y)
            body_rect["h_pt"] = C.emu_to_pt(body_rect["cy"])
            self.record(n, "body_reflow", "ppt/slides/slide%d.xml" % n,
                        {"top": new_y - body_top_shift}, {"top": new_y, "cy": body_rect["cy"]},
                        "title box was extended")

        # --- fit body text -------------------------------------------------------
        ladder = ((self.plan.get("font_ladder") or {}).get("body") or LADDER["body"])
        base = (body_info or {}).get("size")
        if base:
            ladder = [s for s in ladder if s <= base] or [base]
        floor_body = int(floor.get("body") or C.FONT_FLOOR["body"])
        cap_ladder = ((self.plan.get("font_ladder") or {}).get("caption") or LADDER["caption"])
        floor_cap = int(floor.get("caption") or C.FONT_FLOOR["caption"])

        # Resolve figures first: the text box must never be laid out on top of an image.
        fig_items = []
        for fig in figures:
            raw = fig.get("asset") or (self.fig_index.get(fig.get("source")) or {}).get("asset")
            src = resolve_input(self.work, raw) if raw else None
            if src:
                fig_items.append((fig, src))
            else:
                self.action("figure_missing", n, fig.get("source") or "?", "declared figure asset",
                            "asset not resolvable; caption kept, image omitted",
                            [json.dumps(fig, ensure_ascii=False)], "honest degradation")
        img_rids = {i: "rId%d" % (2 + i) for i in range(len(fig_items))}

        gap = 45720
        text_rect = body_rect
        fig_box = None
        cap_plan = {}
        if fig_items and body_rect:
            slots = len(fig_items)
            each_w = max(1, (body_rect["cx"] - gap * (slots - 1)) // max(1, slots))
            # Usable text width/height: minus the body insets the emitted shape will carry,
            # otherwise the reserved strip is ~14pt too wide and ~7pt too short per line.
            cap_w_pt = max(1.0, (C.emu_to_pt(each_w) - 2 * C.emu_to_pt(TEXT_INSET_L)) * SAFE_W)
            base_fig_h = int(body_rect["h_pt"] * 0.45 * C.EMU_PER_PT)
            body_h_emu = int(body_rect["h_pt"] * C.EMU_PER_PT)
            # Size each caption strip from its measured height before fixing the figure band,
            # so the band can grow instead of the caption being squeezed below the floor.
            for i, (fig, src) in enumerate(fig_items):
                cap = plain_text(fig.get("caption_plain") or fig.get("caption") or "")
                csize, need = (cap_ladder[-1] if cap_ladder else floor_cap), 0
                if cap:
                    for cand in cap_ladder:
                        m = C.estimate_wrapped_lines(cap, cand / 100.0, cap_w_pt,
                                                     line_spacing=SAFE_LINE_SPACING)
                        csize, need = cand, int(m["height_pt"] * C.EMU_PER_PT) + 2 * TEXT_INSET_T + gap
                        if need <= int(body_h_emu * 0.2) or cand <= floor_cap:
                            break
                cap_plan[i] = {"size": csize, "need": need, "text": cap}
            cap_max = max([cap_plan[i]["need"] for i in cap_plan] or [0])
            fig_h = min(int(body_h_emu * 0.72),
                        max(base_fig_h, cap_max + int(base_fig_h * 0.35)))
            fig_box = {"x": body_rect["x"],
                       "y": body_rect["y"] + body_h_emu - fig_h,
                       "cx": body_rect["cx"], "cy": fig_h}
            text_h = body_rect["cy"] - fig_h - gap
            if text_h > 0:
                text_rect = dict(body_rect)
                text_rect["cy"] = text_h
                text_rect["h_pt"] = C.emu_to_pt(text_h)

        chosen, metrics, hist = fit_size(bullets, text_rect, ladder, floor_body)
        split_chunks = None
        if chosen is None and bullets:
            chosen = min(floor_body, ladder[-1])
            split_chunks = pack_paragraphs(bullets, text_rect, chosen)
            metrics = dict(metrics or {})
            metrics["action"] = "split"
        elif chosen is not None and bullets and hist and len(hist) > 1:
            metrics["downgraded_from"] = hist[0]["size"]

        # --- figures region ------------------------------------------------------
        fig_rects = []
        if fig_box is not None:
            slots = len(fig_items)
            each_w = max(1, (fig_box["cx"] - gap * (slots - 1)) // max(1, slots))
            for i in range(slots):
                fig_rects.append({"x": fig_box["x"] + i * (each_w + gap), "y": fig_box["y"],
                                  "cx": each_w, "cy": fig_box["cy"]})
            if chosen is not None:
                metrics["figures_share_body"] = True

        # --- assemble shapes -----------------------------------------------------
        tree = _sp_tree()
        added = []
        if title_info is not None or title_key is not None:
            tree.append(_placeholder_sp(self._next_shape_id(), "Title",
                                        title_key[0] if title_key else "title",
                                        title_key[1] if title_key else None,
                                        title_xfrm or (title_info or {}).get("xfrm"),
                                        [{"text": title_text, "level": 0}], title_size, bold=True,
                                        body_pr=(title_info or {}).get("body_pr")))
            added.append("title")
            self.record(n, "placeholder_fill", "ppt/slides/slide%d.xml" % n, None,
                        {"ph": list(title_key) if title_key else None, "sz": title_size,
                         "chars": len(title_text)}, "template_spec.layouts[%s]" % layout_part)

        body_paras = bullets
        if body_paras:
            if body_key is not None:
                tree.append(_placeholder_sp(self._next_shape_id(), "Content Placeholder",
                                           body_key[0], body_key[1], (body_info or {}).get("xfrm"),
                                           body_paras, chosen, body_pr=(body_info or {}).get("body_pr")))
                added.append("body")
                self.record(n, "placeholder_fill", "ppt/slides/slide%d.xml" % n, None,
                            {"ph": list(body_key), "sz": chosen, "bullets": len(body_paras)}, layout_part)
            else:
                sub_rect = body_rect or {"x": 914400, "y": 1600200, "cx": size["cx"] - 1828800,
                                         "cy": size["cy"] - 2286000}
                synth = {"x": sub_rect["x"], "y": sub_rect["y"], "cx": sub_rect["cx"], "cy": sub_rect["cy"]}
                tree.append(_placeholder_sp(self._next_shape_id(), "Substituted Body", None, None,
                                           synth, body_paras, chosen, substituted=True))
                added.append("body_substituted")
                self.action("placeholder_substitute", n, layout_part, "layout has no body placeholder",
                            "explicit text box inside the largest declared rect",
                            ["template_spec.layouts[%s].placeholders" % layout_part],
                            "kept the layout's geometry so overflow stays checkable")
                self.record(n, "placeholder_substitute", "ppt/slides/slide%d.xml" % n, None, synth, layout_part)

        for i, (fig, src) in enumerate(fig_items):
            mem = self.add_media(src)
            nat = C.natural_size(src)
            iw, ih = (nat if nat else (1000, 1000))
            cell = fig_rects[i] if i < len(fig_rects) else {
                "x": body_rect["x"], "y": body_rect["y"], "cx": body_rect["cx"],
                "cy": max(1, int(body_rect["cy"] * 0.4))}
            pl = cap_plan.get(i) or {}
            cap = pl.get("text") or plain_text(fig.get("caption") or "")
            csize = pl.get("size") or (cap_ladder[-1] if cap_ladder else floor_cap)
            cap_h = pl.get("need") or 0
            cap_metrics = None
            if cap:
                cap_metrics = C.estimate_wrapped_lines(
                    cap, csize / 100.0,
                    max(1.0, (C.emu_to_pt(cell["cx"]) - 2 * C.emu_to_pt(TEXT_INSET_L)) * SAFE_W),
                    line_spacing=SAFE_LINE_SPACING)
            img_cell = dict(cell)
            if cap_h:
                img_cell["cy"] = max(1, cell["cy"] - cap_h - gap)
            placed = contain_rect(img_cell, iw, ih)
            tree.append(_picture_sp(self._next_shape_id(), fig.get("source") or "image",
                                    img_rids[i], placed, descr=cap or None))
            added.append("figure:%s" % fig.get("source"))
            self.action("figure_fit", n, mem["part"], "natural %dx%d px" % (iw, ih),
                        "placed %dx%d EMU scale=%s aspect_delta=%s" % (placed["cx"], placed["cy"],
                                                                       placed["scale"], placed["aspect_delta"]),
                        [json.dumps({"cell": img_cell, "natural": [iw, ih]})],
                        "contain fit; a single scale factor preserves the aspect ratio")
            self.record(n, "figure_fit", mem["part"], None, placed, "natural=%dx%d" % (iw, ih))
            if cap:
                cap_rect = {"x": cell["x"], "y": cell["y"] + img_cell["cy"] + gap,
                            "cx": cell["cx"], "cy": max(1, cap_h)}
                tree.append(_placeholder_sp(self._next_shape_id(), "Figure Caption", None, None,
                                            cap_rect, [{"text": cap, "level": 0}], csize, substituted=True))
                added.append("caption:%s" % fig.get("source"))
                self.record(n, "caption", "ppt/slides/slide%d.xml" % n, None,
                            {"sz": csize, "cy": cap_rect["cy"],
                             "lines": (cap_metrics or {}).get("lines")},
                            "template_layout geometry; caption normalised to one paragraph")

        xml = slide_xml(tree)
        rels = {"rId1": {"type": C.RT["slideLayout"], "target": "../slideLayouts/%s" % os.path.basename(layout_part)}}
        for i, (fig, src) in enumerate(fig_items):
            mem = self.add_media(src)
            rels[img_rids[i]] = {"type": C.RT["image"], "target": mem["part"].replace("ppt/", "../")}

        return {"xml": xml, "rels": rels, "layout": layout_part, "meta": meta,
                "added": added, "metrics": metrics, "overflow_history": hist,
                "title_size": title_size, "body_size": chosen,
                "title_ph": list(title_key) if title_key else None,
                "body_ph": list(body_key) if body_key else None,
                "figures": [f.get("source") for f in figures],
                "notes": plan_slide.get("notes"),
                "claims": plan_slide.get("claims") or [],
                "evidence": plan_slide.get("evidence") or [],
                "chunks": split_chunks,
                "body_rect": body_rect,
                "planning": {"role": plan_slide.get("role"), "layout": layout_part,
                             "title_size": title_size, "title_metrics": title_metrics,
                             "body_size": chosen, "body_metrics": metrics,
                             "title_xfrm": title_xfrm, "body_rect": body_rect,
                             "title_box_extended": bool(body_top_shift),
                             "cap_plan": {str(k): v for k, v in (cap_plan or {}).items()}}}

    # ---- whole deck --------------------------------------------------------------
    def run(self):
        plan_slides = list(self.plan.get("slides") or [])
        if not plan_slides:
            return {"ok": False, "detail": "slides_plan.json has no slides"}
        out_slides = []
        n = 0
        for ps in plan_slides:
            n += 1
            ps = dict(ps)
            ps["n"] = n
            built = self.build_slide(ps)
            out_slides.append(built)
            if built.get("chunks") and len(built["chunks"]) > 1:
                self.action("text_split", n, "ppt/slides/slide%d.xml" % n,
                            "%d bullets at floor %d" % (len(ps.get("bullets") or []), built["body_size"]),
                            "%d continuation slide(s); never below the readable floor" % (len(built["chunks"]) - 1),
                            [json.dumps({"chunks": [len(c) for c in built["chunks"]]})],
                            "splitting is preferred over unreadable font sizes")
                for chunk in built["chunks"][1:]:
                    n += 1
                    cont = dict(ps)
                    cont["n"] = n
                    cont["figures"] = []
                    built_c = self.build_slide(cont, continuation_of=ps.get("title") or "",
                                               paragraphs=chunk)
                    out_slides.append(built_c)

        # notes
        res = self.assemble(out_slides, plan_slides)
        res["slide_metrics"] = [dict((b.get("planning") or {}), slide=i + 1)
                                for i, b in enumerate(out_slides)]
        return res

    def ensure_notes_master(self):
        if self.tpl.notes_masters:
            self.notes_master = self.tpl.notes_masters[0]
            return self.notes_master
        self.notes_master = "ppt/notesMasters/notesMaster1.xml"
        self.action("notes_master_added", None, self.notes_master,
                    "template has no notesMaster", "synthesised minimal notesMaster + theme rel",
                    ["template zip entry list"], "required for real speaker notes")
        return self.notes_master

    def assemble(self, out_slides, plan_slides):
        tpl = self.tpl
        parts = {}
        # Copy the template package wholesale and then drop only what we deliberately rebuild.
        # Anything else (printerSettings, embeddings, charts, media, customXml, ...) is kept so
        # the result cannot end up with a dangling relationship — that makes renderers refuse
        # the whole document, and template fidelity is exactly what the task asks for.
        skip_prefixes = ("ppt/slides/", "ppt/notesSlides/")
        skip_exact = {"[Content_Types].xml", "ppt/presentation.xml",
                      "ppt/_rels/presentation.xml.rels"}
        for name, data in tpl.parts.items():
            if name in skip_exact or name.startswith(skip_prefixes):
                continue
            parts[name] = data

        # presentation.xml
        pres = ET.fromstring(tpl.presentation)
        # strip the template's own slides
        for tag in ("p:sldIdLst",):
            node = pres.find(tag, C.NS)
            if node is not None:
                pres.remove(node)
        sld_lst = ET.Element(q("p:sldIdLst"))
        pres_rels = {rid: dict(r) for rid, r in self._parse_rels(tpl.pres_rels).items()}
        max_id = 256
        for rid in list(pres_rels):
            if pres_rels[rid]["type"].endswith("/slide"):
                del pres_rels[rid]
        def next_rid():
            nonlocal max_rid
            max_rid += 1
            return "rId%d" % max_rid
        max_rid = 0
        for rid in pres_rels:
            m = re.match(r"rId(\d+)$", rid)
            if m:
                max_rid = max(max_rid, int(m.group(1)))

        slide_entries = []
        for i, built in enumerate(out_slides, start=1):
            part = "ppt/slides/slide%d.xml" % i
            parts[part] = built["xml"]
            parts["ppt/slides/_rels/slide%d.xml.rels" % i] = self._rels_bytes(built["rels"])
            rid = next_rid()
            pres_rels[rid] = {"type": C.RT["slide"], "target": "slides/slide%d.xml" % i}
            sid = ET.SubElement(sld_lst, q("p:sldId"))
            sid.set("id", str(max_id + i))
            sid.set(q("r:id"), rid)
            slide_entries.append({"slide": i, "part": part, "layout": built["layout"],
                                  "title_ph": built["title_ph"], "body_ph": built["body_ph"]})

        # notes slides
        self.notes_parts = {}
        notes_master_added = None
        any_notes = any((b.get("notes") or "").strip() for b in out_slides)
        if any_notes:
            nm = self.ensure_notes_master()
            if nm not in parts:
                parts[nm] = NOTES_MASTER
                theme = tpl.themes[0] if tpl.themes else None
                rels = {}
                if theme:
                    rels["rId1"] = {"type": C.RT["theme"], "target": "../theme/%s" % os.path.basename(theme)}
                parts["ppt/notesMasters/_rels/notesMaster1.xml.rels"] = self._rels_bytes(rels)
                notes_master_added = nm
            idx = 0
            for i, built in enumerate(out_slides, start=1):
                text = (built.get("notes") or "").strip()
                if not text:
                    continue
                idx += 1
                npart = "ppt/notesSlides/notesSlide%d.xml" % idx
                parts[npart] = notes_xml([text])
                nrels = {"rId1": {"type": C.RT["notesMaster"],
                                  "target": "../notesMasters/%s" % os.path.basename(nm)},
                         "rId2": {"type": C.RT["slide"], "target": "../slides/slide%d.xml" % i}}
                parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % idx] = self._rels_bytes(nrels)
                # link notes from the slide
                srel_part = "ppt/slides/_rels/slide%d.xml.rels" % i
                srels = self._parse_rels(parts[srel_part])
                nid = "rId%d" % (max(int(k[3:]) for k in srels) + 1)
                srels[nid] = {"type": C.RT["notesSlide"], "target": "../notesSlides/notesSlide%d.xml" % idx}
                parts[srel_part] = self._rels_bytes(srels)
                self.action("notes", i, npart, None, text[:80],
                            ["plan.slides[%d].notes" % (i - 1)], "speaker notes written as a real notesSlide")
                self.record(i, "notes", npart, None, {"chars": len(text)}, "plan notes")

        # notes master into presentation
        if notes_master_added:
            nm_lst = pres.find("p:notesMasterIdLst", C.NS)
            if nm_lst is None:
                nm_lst = ET.Element(q("p:notesMasterIdLst"))
                insert_at = 0
                for i, child in enumerate(list(pres)):
                    if C.localname(child.tag) in ("sldMasterIdLst",):
                        insert_at = i + 1
                pres.insert(insert_at, nm_lst)
            rid = next_rid()
            pres_rels[rid] = {"type": C.RT["notesMaster"], "target": "notesMasters/notesMaster1.xml"}
            nid = ET.SubElement(nm_lst, q("p:notesMasterId"))
            nid.set(q("r:id"), rid)
            if pres.find("p:notesSz", C.NS) is None:
                ns = ET.Element(q("p:notesSz"))
                ns.set("cx", "6858000")
                ns.set("cy", "9144000")
                pres.append(ns)

        # insert sldIdLst before sldSz
        sld_sz = pres.find("p:sldSz", C.NS)
        if sld_sz is not None:
            pres.insert(list(pres).index(sld_sz), sld_lst)
        else:
            pres.append(sld_lst)

        parts["ppt/presentation.xml"] = C.xml_bytes(pres)
        parts["ppt/_rels/presentation.xml.rels"] = self._rels_bytes(pres_rels)

        # content types
        parts["[Content_Types].xml"] = self._content_types(tpl.content_types, out_slides, parts)

        # media
        for src, mem in self.media.items():
            parts[mem["part"]] = open(src, "rb").read()

        # write
        deck_path = os.path.join(self.output, self.deck_name)
        self._write_zip(deck_path, parts)
        return {"ok": True, "deck": deck_path, "parts": len(parts),
                "slides": len(out_slides), "media": len(self.media),
                "notes_master_added": notes_master_added,
                "slide_entries": slide_entries}

    def _content_types(self, raw, out_slides, parts):
        """Patch [Content_Types].xml as text.

        Re-serialising this part with ElementTree emits an ``ns0:`` prefix, and OOXML
        consumers (verified against LibreOffice) then refuse the whole package with
        "loadComponentFromURL returned an empty reference".  Text surgery keeps the
        template's own prefix convention untouched.
        """
        text = raw.decode("utf-8")
        text = re.sub(r'<Override[^>]*PartName="/ppt/(?:slides|notesSlides)/[^"]*"[^>]*/>', "", text)
        have = set(parts)

        def _keep_override(m):
            pm = re.search(r'PartName="([^"]+)"', m.group(0))
            return m.group(0) if pm and pm.group(1).lstrip("/") in have else ""

        text = re.sub(r"<Override[^>]*/>", _keep_override, text)

        add = []
        for i in range(1, len(out_slides) + 1):
            add.append('<Override PartName="/ppt/slides/slide%d.xml" ContentType="%s"/>'
                       % (i, CT["sld"]))
        for name in sorted(parts):
            if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", name):
                add.append('<Override PartName="/%s" ContentType="%s"/>' % (name, CT["notes"]))
        if "ppt/notesMasters/notesMaster1.xml" in parts and \
                "/ppt/notesMasters/notesMaster1.xml" not in text:
            add.append('<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="%s"/>'
                       % CT["notesMaster"])
        declared = {d.lower() for d in re.findall(r'<Default[^>]*Extension="([^"]+)"', text)}
        for ext, ctype in IMG_EXT.items():
            if ext not in declared:
                add.append('<Default Extension="%s" ContentType="%s"/>' % (ext, ctype))
        if add:
            idx = text.rindex("</Types>") if "</Types>" in text else text.rindex("</")
            text = text[:idx] + "".join(add) + text[idx:]
        return text.encode("utf-8")

    @staticmethod
    def _parse_rels(raw):
        out = {}
        if not raw:
            return out
        root = ET.fromstring(raw)
        for r in root.findall("{%s}Relationship" % PKG_REL):
            out[r.get("Id")] = {"type": r.get("Type"), "target": r.get("Target"),
                                "mode": r.get("TargetMode")}
        return out

    @staticmethod
    def _rels_bytes(rels):
        root = ET.Element("{%s}Relationships" % PKG_REL)
        for rid, spec in rels.items():
            e = ET.SubElement(root, "{%s}Relationship" % PKG_REL)
            e.set("Id", rid)
            e.set("Type", spec["type"])
            e.set("Target", spec["target"])
            if spec.get("mode"):
                e.set("TargetMode", spec["mode"])
        return C.xml_bytes(root)

    @staticmethod
    def _write_zip(path, parts):
        C.ensure_dir(os.path.dirname(os.path.abspath(path)))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for name in sorted(parts):
                zi = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = 0o600 << 16
                z.writestr(zi, parts[name])


# ======================================================================================
# 5. entry point
# ======================================================================================

def resolve_input(work_dir, rel):
    """Find a staged input by relative or absolute path (discover stages them under <work>/inputs)."""
    if not rel:
        return None
    if os.path.isfile(rel):
        return rel
    cands = [os.path.join(work_dir, rel), os.path.join(work_dir, "inputs", rel),
             os.path.join(work_dir, "_inputs", rel), os.path.join(os.path.dirname(work_dir), rel)]
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def adapt(work_dir, output_dir):
    manifest = C.read_json(os.path.join(work_dir, "case_manifest.json"), {}) or {}
    spec = C.read_json(os.path.join(work_dir, "template_spec.json"), {}) or {}
    inv = C.read_json(os.path.join(work_dir, "source_inventory.json"), {}) or {}
    plan = C.read_json(os.path.join(work_dir, "slides_plan.json"), {}) or {}
    detected = manifest.get("detected") or {}
    tpl_rel = None
    tpl_abs = None
    for cand in ([plan.get("template_path"), spec.get("template_path")]
                 + list(detected.get("template") or [])
                 + [i.get("rel") for i in (manifest.get("inputs") or [])
                    if i.get("kind") in ("pptx", "potx", "ppt")]):
        got = resolve_input(work_dir, cand)
        if got:
            tpl_rel, tpl_abs = cand, got
            break
    if not tpl_abs:
        return {"ok": False, "detail": "no readable template in template_spec/plan/case_manifest inputs"}
    try:
        with zipfile.ZipFile(tpl_abs) as z:
            bad = z.testzip()
        if bad:
            return {"ok": False, "detail": "template zip corrupt at %s" % bad}
    except zipfile.BadZipFile as exc:
        return {"ok": False, "detail": "template is not a valid zip: %s" % exc}

    deck_name = plan.get("deck_name") or C.DECK_DEFAULT_NAME
    ad = Adapter(work_dir, output_dir, tpl_abs, spec, inv, plan, deck_name)
    res = ad.run()
    if not res.get("ok"):
        return res

    counts = {"slides": res["slides"], "placeholders_filled": 0, "substitutions": 0,
              "figure_fits": 0, "font_downgrades": 0, "splits": 0, "notes": 0,
              "figure_missing": 0, "notes_master_added": 1 if res.get("notes_master_added") else 0}
    for a in ad.actions:
        kind = a["kind"]
        if kind in ("placeholder_fill",):
            counts["placeholders_filled"] += 1
        elif kind in ("placeholder_substitute",):
            counts["substitutions"] += 1
        elif kind == "figure_fit":
            counts["figure_fits"] += 1
        elif kind == "text_split":
            counts["splits"] += 1
        elif kind == "notes":
            counts["notes"] += 1
        elif kind == "figure_missing":
            counts["figure_missing"] += 1
    # placeholder_fill is recorded (change ledger) rather than actioned above
    counts["placeholders_filled"] = sum(1 for r in ad.records if r["kind"] == "placeholder_fill")
    counts["notes"] = sum(1 for r in ad.records if r["kind"] == "notes") or counts["notes"]
    for s in res["slide_entries"]:
        pass

    actions = {"ok": True, "deck": deck_name, "deck_sha256": C.sha256_file(res["deck"]),
               "actions": ad.actions, "counts": counts,
               "media": [{"source": s, "part": m["part"], "sha256": m["sha256"], "bytes": m["bytes"]}
                         for s, m in ad.media.items()],
               "notes_master_added": res["notes_master_added"],
               "slide_entries": res["slide_entries"],
               "slide_metrics": res.get("slide_metrics")}
    C.write_json(os.path.join(work_dir, "adaptation_actions.json"), actions)
    with open(os.path.join(work_dir, "change_records.jsonl"), "w", encoding="utf-8") as fh:
        for r in ad.records:
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")
    return {"ok": True, "deck": res["deck"], "actions": actions, "counts": counts,
            "notes_master_added": res["notes_master_added"]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    res = adapt(args.work, args.output)
    print(json.dumps({k: v for k, v in res.items() if k != "actions"}, ensure_ascii=False, indent=2))
    return 0 if res.get("ok") else 3


if __name__ == "__main__":
    raise SystemExit(main())
