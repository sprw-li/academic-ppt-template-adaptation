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


def derive_plan(work):
    """Evidence-only conservative plan: roles already confirmed by discover, no invention."""
    man = _j(work, "case_manifest.json", {}) or {}
    spec = _j(work, "template_spec.json", {}) or {}
    inv = _j(work, "source_inventory.json", {}) or {}
    constraints = man.get("constraints") or {}
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

    # unit ids -> locator for the evidence map
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

    cap = _int_from_constraints(constraints, ["slide_count", "target_slide_count", "slides",
                                              "num_slides", "pages", "max_slides"]) or 16
    per_slide = 5

    slides = []
    # 1 cover
    slides.append({"role": "title", "title": title_text,
                   "subtitle": (inv.get("source_path") and os.path.basename(inv["source_path"])) or "",
                   "claims": ["研究主线来自源材料"], "evidence": [u.get("id") for u in units[:2]],
                   "notes": _notes_for([u for u in units[:2]], loc), **layout_fields("title")})
    # 2 outline
    outline = [{"text": _plain(s.get("title") or ""), "level": 0}
               for s in sections if (s.get("title") or "").strip()][:8]
    if outline:
        slides.append({"role": "outline", "title": "Outline", "bullets": outline,
                       "claims": ["页序对应源材料章节"], "evidence": [s.get("id") for s in sections[:8]],
                       "notes": _notes_for_sections(sections[:8]), **layout_fields("content")})
    # 3 per-section content
    for s in sections:
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
                           "claims": [(s.get("title") or "").strip()],
                           "evidence": [u.get("id") for u in chunk],
                           "notes": _notes_for(chunk, loc), **layout_fields("content")})
    # 4 figure slides
    for f in figures:
        cap_text = _plain(f.get("caption_plain") or f.get("caption") or f.get("label") or "Figure")
        slides.append({"role": "figure", "title": (f.get("label") or "Figure"),
                       "bullets": [], "figures": [{"source": f.get("id"), "asset": f.get("asset"),
                                                   "caption": cap_text, "aspect_lock": True,
                                                   "label": f.get("label")}],
                       "claims": [cap_text], "evidence": [f.get("id")],
                       "notes": _notes_for([{"id": f.get("id"), "locator": f.get("source_ref") or f.get("id")}], loc),
                       **layout_fields("content")})
    # 5 tables
    for t in tables:
        cap_text = _plain(t.get("caption_plain") or t.get("caption") or t.get("label") or "Table")
        slides.append({"role": "table", "title": (t.get("label") or "Table"),
                       "bullets": [{"text": cap_text[:180], "level": 0}],
                       "claims": [cap_text], "evidence": [t.get("id")],
                       "notes": _notes_for([{"id": t.get("id"), "locator": t.get("source_ref") or t.get("id")}], loc),
                       **layout_fields("content")})
    # 6 conclusion (last section's units, if not already the tail)
    if sections:
        last = sections[-1]
        lu = by_section.get(last.get("id")) or []
        if lu and (not slides or slides[-1].get("evidence") != [u.get("id") for u in lu[:per_slide]]):
            slides.append({"role": "conclusion", "title": "Conclusion",
                           "bullets": [{"text": _clip(u.get("text_plain") or u.get("text") or ""),
                                        "level": 0, "evidence": [u.get("id")]} for u in lu[:per_slide]],
                           "claims": ["结论对应源材料"], "evidence": [u.get("id") for u in lu[:per_slide]],
                           "notes": _notes_for(lu[:per_slide], loc), **layout_fields("content")})

    if len(slides) > cap:
        head = slides[:max(1, cap - 1)]
        tail = slides[-1]
        if tail in head:
            slides = head
        else:
            slides = head + [tail]
    for i, s in enumerate(slides, 1):
        s["n"] = i

    deck_name = C.DECK_DEFAULT_NAME
    for k in ("deck_name", "output_name", "primary_output", "deliverable_name"):
        v = constraints.get(k) if isinstance(constraints, dict) else None
        if isinstance(v, str) and v.lower().endswith(".pptx"):
            deck_name = os.path.basename(v)
            break

    plan = {
        "plan_version": "1", "example": False,
        "source_path": inv.get("source_path"),
        "template_path": spec.get("template_path") or (man.get("detected") or {}).get("template", [None])[0],
        "start_deck": (man.get("detected") or {}).get("start_deck", [None])[0] if (man.get("detected") or {}).get("start_deck") else None,
        "deck_name": deck_name,
        "language": "zh" if _looks_cjk(title_text) else "en",
        "audience": "academic audience (from case materials)",
        "time_minutes": None,
        "target_slide_count": len(slides),
        "confirmed_roles": True,
        "slides": slides,
        "font_ladder": {"title": [3200, 2800, 2400], "body": [2000, 1800, 1600, 1400],
                        "caption": [1400, 1200],
                        "floor": {"body": C.FONT_FLOOR["body"], "caption": C.FONT_FLOOR["caption"],
                                  "title": C.FONT_FLOOR["title"]}},
        "numbering": {"mode": "preserve_source",
                      "map": ["%s=%s" % (f.get("id"), f.get("label")) for f in figures if f.get("label")]},
        "contract_values": _contract_values(man),
        "plan_source": "auto_derived",
        "notes": ["conservative plan derived only from discovered evidence; "
                  "roles/layouts/counts come from case_manifest.json, template_spec.json and "
                  "source_inventory.json — nothing was invented"],
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


def _notes_for(units, loc):
    """Speaker notes that only restate source locations — never invent content."""
    refs = []
    for u in units or []:
        if not u:
            continue
        r = loc.get(u.get("id")) or u.get("locator") or u.get("id")
        if r:
            refs.append(str(r))
    return ("来源：" + "；".join(refs)) if refs else ""


def _notes_for_sections(sections):
    names = [str(s.get("title") or s.get("id")) for s in (sections or []) if s]
    return ("章节顺序：" + "；".join(names)) if names else ""


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
