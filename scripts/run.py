#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run.py — orchestration entry point for the academic-PPT template adaptation Skill.

    python run.py prepare --input /input --work /app/work --output /app/output
    python run.py finish  --work /app/work --output /app/output
    python run.py all     --input /input --work /app/work --output /app/output

`prepare` only reconnoitres (never lays out slides); `finish` adapts, audits, renders and
finalises; `all` is the zero-intervention path (prepare -> auto-derived plan -> finish) for
runs under a tight request budget.  Exit codes follow work/CONTRACT.md:

    0  /app/output/result.json exists with status ok|partial
    2  bad arguments or unreadable /input
    3  status needs_human (an honest result.json is still written)

Owner: lead.  Pure standard library, offline, POSIX/Windows safe.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import common as C

PLAN = "slides_plan.json"
RESOLUTION = "plan_resolution.json"
PIPELINE = "pipeline_report.json"
AGENT_PLAN = "slides_plan.agent.json"

EVIDENCE_FILES = ["fidelity_report.json", "format_audit.json", "render_report.json",
                  "change_records.jsonl", "sources.json", "adaptation_actions.json",
                  "plan_resolution.json", "template_spec.json", "source_inventory.json",
                  "case_manifest.json", "pipeline_report.json"]


def _j(work, name, default=None):
    return C.read_json(os.path.join(work, name), default)


def _steps():
    return []


def _pipeline(stage, started, steps, status, outputs=None, stages=None):
    rec = {"stage": stage, "status": status, "started": started, "ended": C.now_iso(),
           "wall_s": round(time.time() - started, 3), "steps": steps,
           "outputs": outputs or []}
    if stages:
        rec["stages"] = stages
    return rec


# ======================================================================================
# plan resolution / auto derivation
# ======================================================================================

def _layout_hints(spec):
    """(title_layout, content_layout, blank_layout) as (name, part) from the real template."""
    layouts = spec.get("layouts") or []
    title = content = blank = None
    for L in layouts:
        name = (L.get("name") or "").strip()
        part = L.get("part")
        phs = L.get("placeholders") or []
        types = set()
        for p in phs:
            t = p.get("ph_type") or p.get("type") or "body"
            types.add(t)
        ltype = (L.get("type") or "").lower()
        if title is None and (ltype in ("title", "titleonly") or "ctrTitle" in types or "title" in ltype):
            title = (name, part)
        if content is None and (types & {"body", "obj"}):
            content = (name, part)
        if blank is None and not types:
            blank = (name, part)
    return title or (None, None), content or (None, None), blank or (None, None)


def _int_from_constraints(constraints, keys):
    if not isinstance(constraints, dict):
        return None
    for k in keys:
        v = constraints.get(k)
        if isinstance(v, int) and v > 0:
            return v
        if isinstance(v, list) and v and all(isinstance(x, int) for x in v):
            return max(v)
    # any key mentioning slide/page with an int value
    for k, v in constraints.items():
        if isinstance(v, int) and v > 1 and ("slide" in k.lower() or "page" in k.lower()):
            return v
    return None


def _constraints(man):
    """Read the frozen case constraints and record where every value came from (v1.3 §A2)."""
    c = man.get("constraints") or {}
    if not isinstance(c, dict):
        c = {}
    applied = {}

    def take(key):
        if key in c and c[key] is not None:
            applied[key] = {"value": c[key], "source": "case.json:constraints.%s" % key}
            return c[key]
        return None

    return {
        "audience": take("audience"),
        "time_minutes": take("time_minutes"),
        "target_slide_count": take("target_slide_count"),
        "font_size_floor_pt": take("font_size_floor_pt"),
        "caption_numbering": take("caption_numbering"),
        "language": take("language"),
        "aspect": take("aspect"),
        "masters_min": take("masters_min"),
        "layouts_min": take("layouts_min"),
    }, applied


def _floor_from_constraints(fc):
    """constraints.font_size_floor_pt is in points; our ladder floor is in 1/100 pt."""
    floor = dict(C.FONT_FLOOR)
    if isinstance(fc, dict):
        for role in ("body", "caption", "title", "table"):
            v = fc.get(role)
            if isinstance(v, (int, float)) and v > 0:
                floor[role] = int(round(float(v) * 100))
    return floor


def _start_deck_slides(sd):
    """Plan entries derived from the organizer-provided starting deck.

    v1.3 §A3: the start deck's text is preserved (re-laid-out and possibly re-split, never
    reworded).  Pictures are referenced by their part name inside the start deck so the
    adapter can copy the real bytes.
    """
    out = []
    for det in (sd.get("slide_details") or []):
        units = det.get("text_units") or []
        title = (det.get("title") or "").strip()
        if not title:
            for u in units:
                if (u.get("where") or "") == "title" and (u.get("text") or "").strip():
                    title = (u.get("text") or "").strip()
                    break
        bullets = []
        for u in units:
            t = (u.get("text") or "").strip()
            if not t or (u.get("where") or "") == "title" or t == title:
                continue
            bullets.append({"text": _clip(t), "level": int(u.get("level") or 0), "evidence": []})
        figs = []
        for pic in (det.get("pictures") or []):
            figs.append({"source_deck_media": pic.get("rel"), "sha256": pic.get("sha256"),
                         "aspect_lock": True, "caption": None})
        tables_here = []
        for t in (det.get("tables") or []):
            tables_here.append({"source": None, "caption": None, "header": [], "rows": [],
                                "evidence": [], "start_deck_table": t})
        entry = {
            "role": "start_deck",
            "title": _clip(title, 160),
            "bullets": bullets,
            "figures": figs,
            "tables": tables_here,
            "source_slide": det.get("index"),
            "preserve_start_text": True,
            "claims": [],
            "evidence": [],
            "start_deck_part": det.get("part"),
            "notes_evidence": ["start_deck:%s" % det.get("part")],
        }
        if det.get("layout_name"):
            entry["start_layout_name"] = det["layout_name"]
        if det.get("placeholder_signature"):
            entry["placeholder_signature"] = det["placeholder_signature"]
        if det.get("placeholder_types"):
            entry["placeholder_types"] = det["placeholder_types"]
        if det.get("layout_probe"):
            entry["start_layout_probe"] = det["layout_probe"]
        out.append(entry)
    return out


def _covered_by_start(title, start_slides):
    """Is this section already represented by a start-deck slide?"""
    key = C.norm_key(_plain(title or ""))
    if not key:
        return False
    for s in start_slides:
        other = C.norm_key(_plain(s.get("title") or ""))
        if other and (key in other or other in key):
            return True
        for b in (s.get("bullets") or []):
            bt = C.norm_key(_plain(b.get("text") or ""))
            if bt and key in bt:
                return True
    return False


def derive_plan(work):
    """Evidence-only conservative plan (v1.3): honours start deck + frozen constraints."""
    man = _j(work, "case_manifest.json", {}) or {}
    spec = _j(work, "template_spec.json", {}) or {}
    inv = _j(work, "source_inventory.json", {}) or {}
    sd = _j(work, "start_deck_spec.json", {}) or {}
    cons, applied = _constraints(man)
    sections = inv.get("sections") or []
    units = inv.get("units") or []
    figures = inv.get("figures") or []
    tables = inv.get("tables") or []

    title_layout, content_layout, _ = _layout_hints(spec)
    if title_layout[0] is None and spec.get("layouts"):
        first = spec["layouts"][0]
        title_layout = (first.get("name"), first.get("part"))
    if content_layout[0] is None:
        content_layout = title_layout

    def layout_fields(kind):
        name, part = (title_layout if kind == "title" else content_layout)
        out = {}
        if name:
            out["layout_name"] = name
        if part:
            out["layout_part"] = part
        return out

    loc = {}
    for u in units:
        loc[u.get("id")] = u.get("locator") or u.get("id")

    by_section = {}
    for u in units:
        key = u.get("section") or (sections[0]["id"] if sections and sections[0].get("id") else "S0")
        by_section.setdefault(key, []).append(u)

    title_text = None
    for s in sections:
        t = (s.get("title") or "").strip()
        if t and len(t) > 2:
            title_text = t
            break
    if not title_text:
        for u in units:
            if u.get("kind") in ("heading", "text") and (u.get("text") or "").strip():
                title_text = u["text"].strip()[:120]
                break
    title_text = title_text or man.get("case_id") or "Presentation"
    title_text = _clip(title_text, 160)

    use_start = bool(sd.get("ok")) and int(sd.get("slide_count") or 0) > 0
    mode = "adapt_start_deck" if use_start else "from_template"
    cap = int(cons["target_slide_count"]) if isinstance(cons.get("target_slide_count"), (int, float)) \
        and cons["target_slide_count"] > 0 else 16
    per_slide = 5

    start_slides = _start_deck_slides(sd) if use_start else []
    slides = list(start_slides)

    # 1 cover (only when the start deck does not already carry one)
    if not (use_start and start_slides and start_slides[0].get("source_slide") == 1):
        slides.append({"role": "title", "title": title_text,
                       "subtitle": (inv.get("source_path") and os.path.basename(inv["source_path"])) or "",
                       "claims": [], "evidence": [u.get("id") for u in units[:2]],
                       "notes_evidence": [u.get("id") for u in units[:2]],
                       **layout_fields("title")})
    # 2 outline
    outline = [{"text": _plain(s.get("title") or ""), "level": 0}
               for s in sections if (s.get("title") or "").strip()][:8]
    if outline:
        slides.append({"role": "outline", "title": "Outline", "bullets": outline,
                       "claims": [], "evidence": [s.get("id") for s in sections[:8]],
                       "notes_evidence": [s.get("id") for s in sections[:8]],
                       **layout_fields("content")})
    # 3 per-section content, skipping sections the start deck already covers
    for s in sections:
        if use_start and _covered_by_start(s.get("title") or "", start_slides):
            continue
        su = by_section.get(s.get("id")) or []
        if not su:
            continue
        body = [u for u in su if u.get("kind") in ("text", "figure_caption", "table_caption", "abstract", "number")]
        if not body:
            body = su
        for i in range(0, len(body), per_slide):
            chunk = body[i:i + per_slide]
            slides.append({"role": "content", "title": _plain(s.get("title") or "Section"),
                           "bullets": [{"text": _clip(u.get("text_plain") or u.get("text") or ""),
                                        "level": 0, "evidence": [u.get("id")]} for u in chunk],
                           "claims": [], "evidence": [u.get("id") for u in chunk],
                           "notes_evidence": [u.get("id") for u in chunk],
                           **layout_fields("content")})
    # 4 figure slides
    for f in figures:
        cap_text = _plain(f.get("caption_plain") or f.get("caption") or f.get("label") or "Figure")
        slides.append({"role": "figure", "title": _plain(f.get("label") or "Figure"),
                       "bullets": [], "figures": [{"source": f.get("id"), "asset": f.get("asset"),
                                                   "caption": cap_text, "aspect_lock": True,
                                                   "label": f.get("label")}],
                       "claims": [], "evidence": [f.get("id")],
                       "notes_evidence": [f.get("id")],
                       **layout_fields("content")})
    # 5 tables: a real editable <a:tbl> only when the source really has cells; otherwise keep
    # the caption honestly and record that the data was not extractable (v1.3 §A3).
    tables_unavailable = []
    for t in tables:
        cap_text = _plain(t.get("caption_plain") or t.get("caption") or t.get("label") or "Table")
        header = [str(x) for x in (t.get("header") or [])]
        rows = [[str(c) for c in r] for r in (t.get("rows") or [])]
        if header or rows:
            slides.append({"role": "table", "title": _plain(t.get("label") or "Table"),
                           "bullets": [], "tables": [{"source": t.get("id"), "caption": cap_text,
                                                      "header": header, "rows": rows,
                                                      "evidence": [t.get("id")],
                                                      "caption_number": t.get("label")}],
                           "claims": [], "evidence": [t.get("id")],
                           "notes_evidence": [t.get("id")],
                           **layout_fields("content")})
        else:
            tables_unavailable.append({"id": t.get("id"), "caption": cap_text,
                                       "data_note": t.get("data_note") or "no real cells extractable",
                                       "evidence": [t.get("id")]})
            slides.append({"role": "table_caption", "title": _plain(t.get("label") or "Table"),
                           "bullets": [{"text": _clip(cap_text), "level": 0, "evidence": [t.get("id")]}],
                           "claims": [], "evidence": [t.get("id")],
                           "notes_evidence": [t.get("id")],
                           **layout_fields("content")})
    # 6 conclusion
    if sections:
        last = sections[-1]
        lu = by_section.get(last.get("id")) or []
        if lu and not (use_start and _covered_by_start(last.get("title") or "", start_slides)):
            slides.append({"role": "conclusion", "title": "Conclusion",
                           "bullets": [{"text": _clip(u.get("text_plain") or u.get("text") or ""),
                                        "level": 0, "evidence": [u.get("id")]} for u in lu[:per_slide]],
                           "claims": [], "evidence": [u.get("id") for u in lu[:per_slide]],
                           "notes_evidence": [u.get("id") for u in lu[:per_slide]],
                           **layout_fields("content")})

    # hard cap on slide count (v1.3 §A5 constraints_honoured).
    # Priority when trimming: start-deck slides (content must be preserved), then the six
    # required page types (title/table/figure/conclusion), and only then redundant body pages.
    if len(slides) > cap:
        must_roles = ("start_deck", "title", "table", "table_caption", "figure", "conclusion")
        must = [s for s in slides if s.get("role") in must_roles]
        optional = [s for s in slides if s.get("role") not in must_roles]
        room = cap - len(must)
        if room < 0:
            # even the required pages exceed the cap: keep order, drop the least critical first
            order = {"start_deck": 0, "title": 1, "conclusion": 2, "table": 3, "figure": 4}
            must.sort(key=lambda s: order.get(s.get("role"), 9))
            keep = set(id(s) for s in must[:cap])
        else:
            keep = set(id(s) for s in must) | set(id(s) for s in optional[:room])
        slides = [s for s in slides if id(s) in keep]
    for i, s in enumerate(slides, 1):
        s["n"] = i

    # narration + evidence (v1.3 §A4) — quoted verbatim from each slide's own evidence
    ev_text = _evidence_text_map(inv, sd)
    for i, s in enumerate(slides):
        nxt = (slides[i + 1].get("title") if i + 1 < len(slides) else "")
        if not (s.get("notes_evidence") or []):
            s["notes_evidence"] = [u.get("id") for u in (s.get("evidence") or []) if u] or \
                                  [s.get("start_deck_part") or "case:title"]
        s["notes"] = _narration(s, ev_text, nxt)

    deck_name = C.DECK_DEFAULT_NAME
    for k in ("deck_name", "output_name", "primary_output", "deliverable_name"):
        v = (man.get("constraints") or {}).get(k) if isinstance(man.get("constraints"), dict) else None
        if isinstance(v, str) and v.lower().endswith(".pptx"):
            deck_name = os.path.basename(v)
            break

    caption_mode = cons.get("caption_numbering")
    plan = {
        "plan_version": "1.3", "example": False,
        "mode": mode,
        "source_path": inv.get("source_path"),
        "template_path": spec.get("template_path") or (man.get("detected") or {}).get("template", [None])[0],
        "start_deck_path": (sd.get("path") if use_start else None),
        "deck_name": deck_name,
        "language": cons.get("language") or ("zh" if _looks_cjk(title_text) else "en"),
        "audience": cons.get("audience"),
        "time_minutes": cons.get("time_minutes"),
        "target_slide_count": cap,
        "constraints_applied": applied,
        "tables_unavailable": tables_unavailable,
        "confirmed_roles": True,
        "slides": slides,
        "font_ladder": {"title": [3200, 2800, 2400], "body": [2000, 1800, 1600, 1400],
                        "caption": [1400, 1200],
                        "floor": _floor_from_constraints(cons.get("font_size_floor_pt"))},
        "numbering": {"mode": caption_mode if caption_mode in ("preserve_source", "relabel") else "preserve_source",
                      "map": ["%s=%s" % (f.get("id"), f.get("label")) for f in figures if f.get("label")]},
        "contract_values": _contract_values(man),
        "plan_source": "auto_derived",
        "notes": ["plan derived only from discovered evidence; mode=%s; start_deck_slides=%d; "
                  "target_slide_count=%s (case constraint)" % (mode, len(start_slides), cap)],
    }
    return plan


def _looks_cjk(text):
    return any("\u2e80" <= ch <= "\u9fff" for ch in (text or ""))


def _plain(s):
    """Prefer discover's plain field, then fall back to stripping markdown locally."""
    try:
        from adapt_pptx import plain_text
        return plain_text(s)
    except Exception:
        import re
        return re.sub(r"\s+", " ", (s or "")).strip()


def _clip(text, limit=200):
    """Keep display text inside ``limit``, cutting at a sentence boundary when possible."""
    text = _plain(text)
    if len(text) <= limit:
        return text
    head = text[:limit]
    for sep in ("。", "！", "？", ". ", "; ", "；"):
        idx = head.rfind(sep)
        if idx > limit * 0.5:
            return head[:idx + len(sep)].rstrip()
    return head.rstrip() + "…"


def _evidence_text_map(inv, sd):
    """unit/section/figure/table id -> its own source text (used to keep notes traceable)."""
    m = {}
    for u in (inv.get("units") or []):
        if u.get("id"):
            m[str(u["id"])] = _plain(u.get("text_plain") or u.get("text") or "")
    for s in (inv.get("sections") or []):
        if s.get("id"):
            m[str(s["id"])] = _plain(s.get("title") or "")
    for f in (inv.get("figures") or []):
        if f.get("id"):
            m[str(f["id"])] = _plain((f.get("label") or "") + " " + (f.get("caption_plain") or f.get("caption") or ""))
    for t in (inv.get("tables") or []):
        if t.get("id"):
            m[str(t["id"])] = _plain((t.get("label") or "") + " " + (t.get("caption_plain") or t.get("caption") or ""))
    for det in (sd.get("slide_details") or []):
        if det.get("part"):
            m["start_deck:%s" % det["part"]] = " ".join(
                _plain(u.get("text") or "") for u in (det.get("text_units") or []))
    return m


def _narration(slide, ev_text=None, next_title=None):
    """A speakable script made **entirely** of text the slide's evidence points at (v1.3 §A4).

    Trade-off, stated plainly: this is a presenter script assembled from the page's own key
    sentences in order, not authored prose.  That is deliberate — an auditor must be able to
    check every content token against the evidence, and authored connective words would be
    unsourced tokens.  The title is prepended only when it actually occurs in the evidence
    text, so the notes stay 100% traceable.  Page order, claims and the evidence map carry
    the narrative logic.
    """
    ev_text = ev_text or {}
    texts = []
    for eid in (slide.get("notes_evidence") or []):
        t = (ev_text.get(str(eid)) or "").strip()
        if t:
            texts.append(_clip(t, 160))
    blob = " ".join(texts)
    key = C.norm_key(blob)
    parts = []
    title = (slide.get("title") or "").strip()
    if title and key and C.norm_key(title) and C.norm_key(title) in key:
        parts.append(title + "。")
    parts.extend(texts[:3])
    if not parts:
        return ""
    return " ".join(parts)




def _contract_values(man):
    """Map original output_contract leaf names onto our real fields (names only, no invention)."""
    contract = man.get("output_contract") or {}
    out = {}
    leaves = _leaves(contract)
    alias = {
        "case_id": "identification.case_id", "task_type": "identification.task_type",
        "status": "delivery.status", "primary_output": "delivery.primary_output",
        "artifacts": "delivery.artifacts", "audit": "delivery.audit",
        "sources": "evidence.sources", "change_records": "evidence.change_records",
        "fidelity": "evidence.fidelity", "format_audit": "evidence.format_audit",
        "render": "evidence.render", "uncertainty": "open_items.uncertainty",
        "limitations": "open_items.limitations", "runtime": "runtime",
    }
    for leaf in leaves:
        key = leaf.rsplit(".", 1)[-1]
        if key in alias:
            out[leaf] = alias[key]
    return out


def _leaves(obj, prefix="", out=None):
    if out is None:
        out = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            _leaves(v, ("%s.%s" % (prefix, k)).lstrip("."), out)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            _leaves(v, "%s[%d]" % (prefix, i), out)
    else:
        out.append(prefix)
    return out


# ======================================================================================
# stages
# ======================================================================================

def write_evidence(work, output):
    dst = os.path.join(output, "evidence")
    C.ensure_dir(dst)
    copied = []
    for name in EVIDENCE_FILES:
        src = os.path.join(work, name)
        if os.path.isfile(src):
            C.safe_copy(src, os.path.join(dst, name))
            copied.append(name)
    return copied


def stage_prepare(input_dir, work, output):
    steps = []
    started = time.time()
    C.ensure_dir(work)
    C.ensure_dir(output)
    try:
        import discover
    except ImportError as exc:
        return {"ok": False, "detail": "discover module missing: %s" % exc, "steps": steps}, 3
    res = discover.discover(input_dir, work, output)
    steps.append({"name": "discover", "status": "ok" if res.get("ok", True) else "failed",
                  "detail": json.dumps({k: v for k, v in (res or {}).items()
                                        if k in ("ok", "case_id", "detail", "notes")}, ensure_ascii=False)[:400]})
    produced = [n for n in ("case_manifest.json", "template_spec.json", "source_inventory.json",
                            "slide_catalog.jsonl", "REVIEW_PACKET.md", "sources.json")
                if os.path.exists(os.path.join(work, n))]
    steps.append({"name": "artifacts", "status": "ok" if produced else "failed", "detail": ",".join(produced)})
    spec = _j(work, "template_spec.json", {}) or {}
    if not spec.get("template_path") and not (spec.get("masters") or spec.get("layouts")):
        steps.append({"name": "template_usable", "status": "failed",
                      "detail": "no readable pptx template in the case materials"})
    rec = _pipeline("prepare", started, steps, "ok" if produced else "failed")
    C.write_json(os.path.join(work, PIPELINE), rec)
    return {"ok": bool(produced), "steps": steps, "produced": produced}, (0 if produced else 3)


def resolve_plan(work):
    """Decide where slides_plan.json comes from; never silently trust a broken plan."""
    path = os.path.join(work, PLAN)
    existing = _j(work, PLAN, None)
    problems = []
    if isinstance(existing, dict) and existing.get("slides"):
        if not existing.get("template_path") and not (existing.get("slides") or [{}])[0].get("layout_part"):
            problems.append("plan has no template_path/layout_part")
        if existing.get("confirmed_roles") is False:
            problems.append("plan declares confirmed_roles=false")
        if not problems:
            C.write_json(os.path.join(work, RESOLUTION),
                         {"plan_source": "agent", "reasons": [], "agent_plan_backup": None,
                          "slides": len(existing["slides"])})
            return existing, {"plan_source": "agent", "reasons": [], "slides": len(existing["slides"])}
    derived = derive_plan(work)
    if isinstance(existing, dict) and existing.get("slides"):
        C.write_json(os.path.join(work, AGENT_PLAN), existing)
        backup = AGENT_PLAN
    else:
        backup = None
    C.write_json(path, derived)
    res = {"plan_source": "auto_derived", "reasons": problems or ["no usable agent plan found"],
           "agent_plan_backup": backup, "slides": len(derived["slides"])}
    C.write_json(os.path.join(work, RESOLUTION), res)
    return derived, res


def stage_finish(work, output, prepare_s=None):
    steps = []
    started = time.time()
    C.ensure_dir(work)
    C.ensure_dir(output)
    plan, resolution = resolve_plan(work)
    steps.append({"name": "plan", "status": "ok",
                  "detail": "plan_source=%s slides=%d" % (resolution["plan_source"], resolution["slides"])})

    missing = []
    for mod in ("adapt_pptx", "verify_pptx", "render_report", "finalize"):
        try:
            __import__(mod)
        except ImportError as exc:
            missing.append("%s (%s)" % (mod, exc))
    if "adapt_pptx" in " ".join(missing):
        return {"ok": False, "detail": "adapter module missing", "steps": steps, "status": "needs_human"}, 3

    import adapt_pptx
    ad = adapt_pptx.adapt(work, output)
    steps.append({"name": "adapt", "status": "ok" if ad.get("ok") else "failed",
                  "detail": json.dumps(ad.get("counts") or ad.get("detail"), ensure_ascii=False)[:300]})
    if not ad.get("ok"):
        _fallback_result(work, output, plan, resolution, "needs_human",
                         "adaptation failed: %s" % ad.get("detail"))
        return {"ok": False, "steps": steps, "status": "needs_human"}, 3

    if "verify_pptx" in " ".join(missing) or "render_report" in " ".join(missing) \
            or "finalize" in " ".join(missing):
        detail = "auditor module missing: %s" % missing
        steps.append({"name": "audit", "status": "failed", "detail": detail})
        write_evidence(work, output)
        _fallback_result(work, output, plan, resolution, "needs_human", detail)
        rec = _pipeline("finish", started, steps, "needs_human", [os.path.basename(ad.get("deck") or "")])
        C.write_json(os.path.join(work, PIPELINE), rec)
        return {"ok": False, "steps": steps, "status": "needs_human", "detail": detail}, 3

    import verify_pptx, render_report, finalize
    fid = verify_pptx.verify(work, output)
    steps.append({"name": "verify", "status": "ok" if (fid or {}).get("ok") else "failed",
                  "detail": "fidelity.ok=%s format=%s" % ((fid or {}).get("ok"),
                                                          (_j(work, "format_audit.json", {}) or {}).get("status"))})
    rnd = render_report.render(work, output)
    steps.append({"name": "render", "status": (rnd or {}).get("status", "skipped"),
                  "detail": ((rnd or {}).get("renderer") or {}).get("name")
                            if isinstance((rnd or {}).get("renderer"), dict) else str((rnd or {}).get("renderer"))})

    rec = _pipeline("finish", started, steps, "ok")
    # The per-stage split must exist *before* finalize reads pipeline_report.json, otherwise
    # result.json.runtime cannot report a real prepare duration.
    rec["stages"] = ([{"stage": "prepare", "wall_s": prepare_s}] if prepare_s is not None else []) + \
        [{"stage": "finish", "wall_s": rec["wall_s"], "measured": "through_audit"}]
    rec["prepare_wall_s"] = prepare_s
    C.write_json(os.path.join(work, PIPELINE), rec)
    finalize.build(work, output)
    val = finalize.validate(work, output)
    copied = write_evidence(work, output)
    steps.append({"name": "finalize", "status": "ok" if (val or {}).get("ok") else "failed",
                  "detail": "evidence=%d validate.ok=%s" % (len(copied), (val or {}).get("ok"))})
    result = _j(output, "result.json", {}) or {}
    status = ((result.get("delivery") or {}).get("status")) or "needs_human"
    rec = _pipeline("finish", started, steps, status, [os.path.basename(ad.get("deck") or "")])
    C.write_json(os.path.join(work, PIPELINE), rec)
    if (val or {}).get("ok") is False and status == "ok":
        status = "partial"
    return {"ok": status in ("ok", "partial"), "steps": steps, "status": status,
            "result": os.path.join(output, "result.json")}, (0 if status in ("ok", "partial") else 3)


def _fallback_result(work, output, plan, resolution, status, detail):
    """Last-resort honest result.json so a needs_human run is never empty."""
    man = _j(work, "case_manifest.json", {}) or {}
    deck_name = (plan or {}).get("deck_name") or C.DECK_DEFAULT_NAME
    deck_present = os.path.isfile(os.path.join(output, deck_name))
    result = {
        "identification": {"case_id": man.get("case_id") or (plan or {}).get("case_id") or "unknown",
                           "task_type": man.get("task_type") or "academic_ppt_template_adaptation",
                           "skill": C.SKILL_NAME, "skill_version": C.SKILL_VERSION},
        # never claim an output that does not exist
        "delivery": {"status": status, "primary_output": deck_name if deck_present else None,
                     "artifacts": [], "audit": {"fidelity_ok": False, "format_status": "skipped",
                                                "render_status": "skipped",
                                                "unfilled_fields": [], "fallback": True}},
        "evidence": {}, "slides": [], "plan": {"plan_source": (resolution or {}).get("plan_source")},
        "open_items": {"uncertainty": [detail],
                       "limitations": ["result.json was written by run.py fallback because the "
                                       "finalizer was unavailable or adaptation failed",
                                       "no editable deck was produced for this case"] if not deck_present
                                      else ["result.json was written by run.py fallback"]},
        "runtime": {"python": sys.version.split()[0], "detail": detail},
        "output_contract": man.get("output_contract") or {},
    }
    # §6.1: even the degraded report keeps the original contract snapshot, a path map for the
    # leaves we can fill, and an explicit list of the leaves we cannot — nothing is silently dropped.
    mapped = _contract_values(man)
    unfilled = [l for l in _leaves(result["output_contract"]) if l not in mapped]
    result["contract_path_map"] = mapped
    result["delivery"]["audit"]["unfilled_fields"] = unfilled
    C.write_json(os.path.join(output, "result.json"), result)


def main(argv=None):
    ap = argparse.ArgumentParser(prog="run.py")
    ap.add_argument("stage", choices=["prepare", "finish", "all"])
    ap.add_argument("--input", default="/input")
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--output", default="/app/output")
    args = ap.parse_args(argv)

    if not os.path.isdir(args.work) and args.stage in ("finish",):
        C.log("work dir missing: %s" % args.work, "fatal")
        return 2
    if args.stage in ("prepare", "all") and not os.path.isdir(args.input):
        C.log("input dir missing: %s" % args.input, "fatal")
        return 2

    if args.stage == "prepare":
        res, code = stage_prepare(args.input, args.work, args.output)
    elif args.stage == "finish":
        res, code = stage_finish(args.work, args.output)
    else:
        t0 = time.time()
        res, code = stage_prepare(args.input, args.work, args.output)
        prepare_s = round(time.time() - t0, 3)
        if code != 0 and not os.path.exists(os.path.join(args.work, "case_manifest.json")):
            C.log("prepare failed hard: %s" % res.get("detail"), "fatal")
            return code
        t1 = time.time()
        res2, code = stage_finish(args.work, args.output, prepare_s=prepare_s)
        finish_s = round(time.time() - t1, 3)
        # final, complete per-stage record (finalize already read the earlier, honest version)
        pres = _j(args.work, PIPELINE, {}) or {}
        pres["stages"] = [{"stage": "prepare", "wall_s": prepare_s},
                          {"stage": "finish", "wall_s": finish_s, "measured": "complete"}]
        pres["prepare_wall_s"] = prepare_s
        pres["finish_wall_s"] = finish_s
        C.write_json(os.path.join(args.work, PIPELINE), pres)
        res = {"prepare": res, "finish": res2,
               "stages": pres["stages"]}
    C.log("stage=%s exit=%d" % (args.stage, code))
    print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
