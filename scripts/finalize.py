#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""finalize.py — assemble /app/output/result.json and validate it honestly.

Frozen interface (CONTRACT §4.1)::

    build(work_dir, output_dir) -> dict
        writes <output>/result.json (+ <output>/slides_plan.json, <output>/README.md)
        and the final <work>/sources.json
    validate(work_dir, output_dir) -> dict
        writes <output>/evidence/validation_report.json ; must contain ok: bool

Hard rules enforced here:

* the original ``output_contract`` is preserved **verbatim** (never renamed, flattened
  or dropped) and mirrored to ``result.json.output_contract``;
* every leaf of that contract gets a dotted ``contract_path_map`` entry pointing at a
  real result field, or is listed in ``delivery.audit.unfilled_fields`` and downgrades
  ``delivery.status`` — nothing is silently invented;
* ``status=ok`` is only allowed when fidelity ok, format pass, render rendered and no
  unfilled fields.  exit 0 never implies ok.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import common  # noqa: E402
import verify_pptx  # noqa: E402  (shared read-only pptx parsing, same directory)

SKILL = common.SKILL_NAME
VERSION = common.SKILL_VERSION
HARD_FAIL_STATUS = "needs_human"

# a statement that is only true *before* the finish stage; it must never be presented as a
# final limitation once a deck has actually been produced
_STAGE_CLAIM_RX = re.compile(
    r"no slide|no fidelity|no format|no render|not produced|no claim|未产出|未生成|未做|"
    r"尚未|尚无|not yet produced|nothing is produced", re.I)


def _split_limitations(items, status, deck_present):
    """Return (final_limitations, prepare_stage_notes).

    Stage-scoped statements ("prepare inspects inputs only; no slide is produced") are
    only kept as final limitations while the delivery really has no deck; on a successful
    run they are demoted to ``open_items.prepare_stage_notes`` with a resolution flag so
    the result never contradicts ``delivery.status``.
    """
    final, staged = [], []
    for item in items or []:
        if not isinstance(item, str) or not item.strip():
            continue
        if _STAGE_CLAIM_RX.search(item):
            staged.append({"prepare_stage_note": item,
                           "resolved_in_finish": bool(deck_present and status in ("ok", "partial"))})
            continue
        final.append(item)
    if not (deck_present and status in ("ok", "partial")):
        final.extend(s["prepare_stage_note"] for s in staged)
    return final, staged


def _j(path, default=None):
    v = common.read_json(path, default=default)
    return v if v is not None else default


def _leaves(obj, prefix=""):
    """Dotted leaf paths of a nested contract. Empty dict/list is itself a leaf."""
    out = []
    if isinstance(obj, dict):
        if not obj:
            out.append(prefix)
            return out
        for k, v in obj.items():
            p = "%s.%s" % (prefix, k) if prefix else str(k)
            out.extend(_leaves(v, p))
    elif isinstance(obj, list):
        out.append(prefix)
    else:
        out.append(prefix)
    return out


def _resolve(obj, dotted):
    cur = obj
    for part in dotted.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return (False, None)
    return (True, cur)


def _artifact(output_dir, rel, kind, **extra):
    path = os.path.join(output_dir, rel)
    if not os.path.isfile(path):
        return None
    entry = {"rel": rel, "sha256": common.sha256_file(path), "bytes": os.path.getsize(path),
             "kind": kind}
    entry.update(extra)
    return entry


def _evidence_entry(work_dir, name, rel_to_output):
    path = os.path.join(work_dir, name)
    if not os.path.isfile(path):
        return None
    return {"rel": rel_to_output, "sha256": common.sha256_file(path),
            "bytes": os.path.getsize(path)}


def _slide_links(deck_path):
    """[{slide, layout}] + masters/layouts/theme actually present in the deck package."""
    info = {"slide_layout_links": [], "masters": [], "layouts": [], "theme": None}
    if not deck_path or not os.path.isfile(deck_path):
        return info
    try:
        with zipfile.ZipFile(deck_path) as z:
            names = z.namelist()
            info["masters"] = sorted(n for n in names
                                     if re.match(r"^ppt/slideMasters/slideMaster\d+\.xml$", n))
            info["layouts"] = sorted(n for n in names
                                     if re.match(r"^ppt/slideLayouts/slideLayout\d+\.xml$", n))
            themes = sorted(n for n in names if re.match(r"^ppt/theme/theme\d+\.xml$", n))
            info["theme"] = themes[0] if themes else None
            slides = sorted((n for n in names if re.match(r"^ppt/slides/slide\d+\.xml$", n)),
                            key=lambda n: int(re.search(r"(\d+)", os.path.basename(n)).group(1)))
            for i, sp in enumerate(slides, 1):
                layout = None
                rel_name = "ppt/slides/_rels/%s.rels" % os.path.basename(sp)
                raw = z.read(rel_name) if rel_name in names else None
                if raw:
                    from xml.etree import ElementTree as ET
                    try:
                        root = ET.fromstring(raw)
                    except ET.ParseError:
                        root = None
                    if root is not None:
                        for rel in root:
                            if (rel.get("Type") or "").endswith("/slideLayout"):
                                tgt = (rel.get("Target") or "").split("/")[-1]
                                layout = "ppt/slideLayouts/" + tgt
                                break
                info["slide_layout_links"].append({"slide": i, "layout": layout})
    except Exception:
        pass
    return info


def _status_of(fidelity_present, format_present, fidelity_ok, format_status,
               render_status, unfilled):
    """CONTRACT v1.2 status semantics.

    ``needs_human`` is reserved for exactly three hard failures:

    * ``fidelity_report.ok == false`` (fabricated / untraceable content);
    * a critical format check failed (``format_audit.status == "fail"``);
    * an auditor module missing (its report was never produced).

    An unavailable renderer is **not** a needs_human condition: the editable deck is
    still usable, only the preview is absent, so the delivery degrades to ``partial``
    (honest, exit 0).  ``render_report.status`` itself stays ``skipped``/``failed``.
    """
    if not fidelity_present or not format_present:
        return HARD_FAIL_STATUS
    if fidelity_ok is False or format_status == "fail":
        return HARD_FAIL_STATUS
    if format_status == "partial" or unfilled:
        return "partial"
    if render_status == "rendered":
        return "ok"
    return "partial"


def build(work_dir: str, output_dir: str) -> dict:
    work_dir = os.path.abspath(work_dir)
    output_dir = os.path.abspath(output_dir)
    common.ensure_dir(output_dir)

    manifest = _j(os.path.join(work_dir, "case_manifest.json"), {}) or {}
    plan = _j(os.path.join(work_dir, "slides_plan.json"), {}) or {}
    resolution = _j(os.path.join(work_dir, "plan_resolution.json"), {}) or {}
    spec = _j(os.path.join(work_dir, "template_spec.json"), {}) or {}
    source_inv = _j(os.path.join(work_dir, "source_inventory.json"), {}) or {}
    fid = _j(os.path.join(work_dir, "fidelity_report.json"), {}) or {}
    fmt = _j(os.path.join(work_dir, "format_audit.json"), {}) or {}
    rnd = _j(os.path.join(work_dir, "render_report.json"), {}) or {}
    pipeline = _j(os.path.join(work_dir, "pipeline_report.json"), {}) or {}
    sources_in = _j(os.path.join(work_dir, "sources.json"), {}) or {}
    changes = common.read_jsonl(os.path.join(work_dir, "change_records.jsonl"))

    deck_path = verify_pptx._find_deck(output_dir, plan)
    deck_rel = None
    if deck_path:
        deck_rel = common.relpath_or_abspath(deck_path, output_dir)

    output_contract = manifest.get("output_contract")
    if output_contract is None:
        output_contract = {}
    preserved = json.loads(json.dumps(output_contract, ensure_ascii=False))

    # ----- artifacts -----
    artifacts = []
    if deck_rel:
        art = _artifact(output_dir, deck_rel, "pptx", editable=True)
        if art:
            artifacts.append(art)
    pdf_art = _artifact(output_dir, "preview/slides.pdf", "pdf", rendered=True)
    if pdf_art and (rnd.get("status") == "rendered"):
        artifacts.append(pdf_art)
    page_arts = []
    for pg in (rnd.get("pages") or []):
        rel = pg.get("rel")
        a = _artifact(output_dir, rel, "png", rendered=True)
        if a:
            page_arts.append(a)
    artifacts.extend(page_arts)

    # ----- delivery / status -----
    fidelity_ok = fid.get("ok")
    format_status = fmt.get("status")
    render_status = rnd.get("status") or "skipped"
    missing_required = [n for n in ("case_manifest.json", "source_inventory.json")
                        if not os.path.isfile(os.path.join(work_dir, n))]
    if deck_path is None:
        missing_required.append("<output>/<deck>.pptx")

    # ----- contract mapping (verbatim contract + dotted path map) -----
    root_fields = {
        "identification": {}, "delivery": {}, "evidence": {}, "slides": [],
        "plan": {}, "template_inheritance": {}, "open_items": {}, "runtime": {},
    }
    placeholder = {
        "identification": {
            "case_id": manifest.get("case_id"),
            "task_type": manifest.get("task_type", "academic_ppt_template_adaptation"),
            "skill": SKILL, "skill_version": VERSION,
        },
        "delivery": {"status": None, "primary_output": deck_rel,
                     "artifacts": artifacts, "audit": {}},
        "evidence": {
            "sources": _evidence_entry(work_dir, "sources.json", "evidence/sources.json"),
            "change_records": {"rel": "evidence/change_records.jsonl", "records": len(changes)},
            "fidelity": _evidence_entry(work_dir, "fidelity_report.json", "evidence/fidelity_report.json"),
            "format_audit": _evidence_entry(work_dir, "format_audit.json", "evidence/format_audit.json"),
            "render": _evidence_entry(work_dir, "render_report.json", "evidence/render_report.json"),
        },
        "slides": [],
        "plan": {}, "template_inheritance": {}, "open_items": {}, "runtime": {},
    }
    # contract_path_map + status are computed *after* every section is populated (below)

    # ----- slides (plan -> result, with locators) -----
    unit_index = {}
    for u in (source_inv.get("units") or []):
        if isinstance(u, dict) and u.get("id"):
            unit_index[u["id"]] = u
    fig_index = {}
    for key in ("figures", "tables"):
        for f in (source_inv.get(key) or []):
            if isinstance(f, dict) and f.get("id"):
                fig_index[f["id"]] = f
    slides_out = []
    for sl in (plan.get("slides") or []):
        if not isinstance(sl, dict):
            continue
        ev = []
        for ref in (sl.get("evidence") or []):
            if not isinstance(ref, str):
                continue
            if ref in unit_index:
                ev.append({"kind": "unit", "ref": ref,
                           "locator": unit_index[ref].get("locator"),
                           "cite": (unit_index[ref].get("text") or "")[:120]})
            elif ref in fig_index:
                kind = "figure" if ref.startswith("F") else "table"
                ev.append({"kind": kind, "ref": ref,
                           "locator": fig_index[ref].get("source_ref"),
                           "cite": (fig_index[ref].get("caption") or "")[:120]})
            else:
                ev.append({"kind": "unit", "ref": ref, "locator": None, "cite": None})
        slides_out.append({"n": sl.get("n"), "role": sl.get("role"), "title": sl.get("title"),
                           "claims": sl.get("claims") or [], "evidence": ev})
    placeholder["slides"] = slides_out

    # ----- plan / inheritance -----
    links = _slide_links(deck_path)
    layout_usage = {}
    for sl in (plan.get("slides") or []):
        if isinstance(sl, dict) and sl.get("layout_name"):
            layout_usage[sl["layout_name"]] = layout_usage.get(sl["layout_name"], 0) + 1
    placeholder["plan"] = {
        "template": plan.get("template_path") or spec.get("template_path"),
        "source": plan.get("source_path") or source_inv.get("source_path"),
        "slide_count": len(slides_out),
        "plan_source": resolution.get("plan_source"),
        "layout_usage": layout_usage,
    }
    placeholder["template_inheritance"] = {
        "masters": links["masters"], "layouts": links["layouts"], "theme": links["theme"],
        "slide_layout_links": links["slide_layout_links"],
    }
    lim_final, lim_staged = _split_limitations(
        (sources_in.get("limitations") or []) + [
            "Local fixture evidence only; hidden platform cases may differ."],
        "partial", bool(deck_rel))   # provisional; recomputed once the real status is known
    placeholder["open_items"] = {
        "uncertainty": (source_inv.get("notes") or []) + (resolution.get("reasons") or [])
        + (plan.get("notes") or []),
        "limitations": lim_final,
        "prepare_stage_notes": lim_staged,
    }
    budget = (manifest.get("budget") or {})
    stages = pipeline.get("stages") or pipeline.get("steps") or []
    wall = pipeline.get("wall_s")
    wall = float(wall) if isinstance(wall, (int, float)) else 0.0

    def _stage_seconds(name):
        for st in stages:
            if not isinstance(st, dict):
                continue
            if str(st.get("stage") or st.get("name") or "").lower() == name:
                v = st.get("wall_s", st.get("seconds"))
                if isinstance(v, (int, float)):
                    return float(v)
        return None

    prep = pipeline.get("prepare_wall_s")
    fin = pipeline.get("finish_wall_s")
    prep = float(prep) if isinstance(prep, (int, float)) else _stage_seconds("prepare")
    fin = float(fin) if isinstance(fin, (int, float)) else _stage_seconds("finish")
    stage_now = str(pipeline.get("stage") or "").lower()
    if prep is None:
        prep = wall if stage_now == "prepare" else 0.0
    if fin is None:
        fin = wall if stage_now in ("finish", "all", "") else 0.0
    if _stage_seconds("prepare") is not None or _stage_seconds("finish") is not None:
        timing = "pipeline_report.stages[]"
    elif pipeline.get("prepare_wall_s") is not None or pipeline.get("finish_wall_s") is not None:
        timing = "pipeline_report.prepare_wall_s/finish_wall_s"
    else:
        timing = "pipeline_report.wall_s (per-stage split unavailable)"
    placeholder["runtime"] = {
        "prepare_s": round(prep, 3),
        "finish_s": round(fin, 3),
        "wall_s": round(wall, 3),
        "timing_source": timing,
        "python": "%d.%d.%d" % sys.version_info[:3],
        "renderer": (rnd.get("renderer") or {}).get("name") or "none",
        "budget": {"model_requests": budget.get("model_requests", 64),
                   "seconds": budget.get("seconds", 900)},
    }

    # ----- contract path map + honest status (all sections are now populated) -----
    placeholder["delivery"]["audit"] = {
        "fidelity_ok": fidelity_ok,
        "format_status": format_status,
        "render_status": render_status,
        "unfilled_fields": [],
        "vacuous_skips": int(fid.get("vacuous_skips", 0)) + int(fmt.get("vacuous_skips", 0)),
    }
    path_map, unfilled = {}, []
    placeholder["delivery"]["status"] = "partial"  # provisional: the status leaf is never null
    for leaf in _leaves(output_contract):
        ok, val = _resolve(placeholder, leaf)
        if ok and val is not None:
            path_map[leaf] = leaf
        else:
            unfilled.append(leaf)
    placeholder["delivery"]["audit"]["unfilled_fields"] = unfilled
    fid_present = os.path.isfile(os.path.join(work_dir, "fidelity_report.json"))
    fmt_present = os.path.isfile(os.path.join(work_dir, "format_audit.json"))
    status = _status_of(fid_present, fmt_present, fidelity_ok, format_status,
                        render_status, unfilled)
    placeholder["delivery"]["status"] = status

    # now that the status is final, make the limitations status-consistent
    lim_final, lim_staged = _split_limitations(
        (sources_in.get("limitations") or []) + [
            "Local fixture evidence only; hidden platform cases may differ."],
        status, bool(deck_rel))
    placeholder["open_items"]["limitations"] = lim_final
    placeholder["open_items"]["prepare_stage_notes"] = lim_staged

    result = {
        "identification": placeholder["identification"],
        "delivery": placeholder["delivery"],
        "evidence": placeholder["evidence"],
        "slides": placeholder["slides"],
        "plan": placeholder["plan"],
        "template_inheritance": placeholder["template_inheritance"],
        "open_items": placeholder["open_items"],
        "runtime": placeholder["runtime"],
        "contract_path_map": path_map,
        "generated_at": common.now_iso(),
        "output_contract": preserved,
    }
    # guard: never let contract_path_map / unfilled be dropped by a later edit
    result["delivery"]["audit"]["unfilled_fields"] = unfilled
    common.write_json(os.path.join(output_dir, "result.json"), result)

    # ----- public copies: slides_plan.json + README.md -----
    if plan:
        common.write_json(os.path.join(output_dir, "slides_plan.json"), plan)
    readme = _readme(result, output_dir, fid, fmt, rnd, unfilled, missing_required)
    common.write_text(os.path.join(output_dir, "README.md"), readme)

    # ----- final sources.json in <work> -----
    sources = dict(sources_in) if sources_in else {}
    sources.setdefault("inputs", [{"rel": i.get("rel"), "kind": i.get("kind"),
                                   "sha256": i.get("sha256")}
                                  for i in (manifest.get("inputs") or []) if isinstance(i, dict)])
    sources["template_sha256"] = spec.get("sha256") or sources.get("template_sha256")
    sources["deck_sha256"] = common.sha256_file(deck_path) if deck_path else None
    sources["skill"] = {"name": SKILL, "version": VERSION,
                        "scripts_sha256": {"common.py": common.sha256_file(
                            os.path.join(_HERE, "common.py"))}}
    sources["toolchain"] = {"python": "%d.%d.%d" % sys.version_info[:3],
                            "renderer": (rnd.get("renderer") or {}).get("name") or "none",
                            "pdf_text": (rnd.get("fonts") or {}).get("tool") or "none"}
    sources["reproduce"] = {
        "prepare": "python run.py prepare --input /input --work /app/work --output /app/output",
        "finish": "python run.py finish --work /app/work --output /app/output",
    }
    sources.setdefault("uncertainty", placeholder["open_items"]["uncertainty"])
    sources.setdefault("limitations", placeholder["open_items"]["limitations"])
    common.write_json(os.path.join(work_dir, "sources.json"), sources)

    return {"ok": status in ("ok", "partial"), "status": status, "result_json":
            os.path.join(output_dir, "result.json"), "unfilled_fields": unfilled,
            "contract_path_map": path_map, "artifacts": artifacts,
            "fidelity_ok": fidelity_ok, "format_status": format_status,
            "render_status": render_status}


def _readme(result, output_dir, fid, fmt, rnd, unfilled, missing_required):
    d = result["delivery"]
    lines = [
        "# Deliverable — academic PPT template adaptation",
        "",
        "Status: **%s**" % d["status"],
        "",
        "## Artifacts",
        "",
    ]
    for a in d["artifacts"]:
        lines.append("- `%s` (%s, %d bytes, sha256 `%s`)" % (a["rel"], a["kind"],
                                                             a["bytes"], a["sha256"][:16]))
    lines += ["", "## Audit summary", "",
              "- fidelity ok: `%s`" % d["audit"]["fidelity_ok"],
              "- format status: `%s`" % d["audit"]["format_status"],
              "- render status: `%s`" % d["audit"]["render_status"],
              "- vacuous skips: `%s`" % d["audit"].get("vacuous_skips"),
              "- unchecked checks: `%d` (see evidence/format_audit.json `unchecked`)"
              % len((fmt or {}).get("unchecked") or []),
              ""]
    if unfilled:
        lines += ["## Unfilled contract fields", ""]
        lines += ["- `%s`" % f for f in unfilled]
        lines += [""]
    if missing_required:
        lines += ["## Missing inputs", ""]
        lines += ["- `%s`" % m for m in missing_required]
        lines += [""]
    lines += ["## How to reproduce", "",
              "```bash",
              "python run.py all --input /input --work /app/work --output /app/output",
              "python -B work/verify/check.py   # local acceptance (see work/verify/test-summary.json)",
              "```",
              "",
              "## Limits / uncertainty", ""]
    for u in result["open_items"]["uncertainty"][:10]:
        lines.append("- %s" % u)
    for l in result["open_items"]["limitations"][:10]:
        lines.append("- %s" % l)
    lines.append("")
    return "\n".join(lines)


def validate(work_dir: str, output_dir: str) -> dict:
    """Independent check of /app/output/result.json. Writes evidence/validation_report.json."""
    work_dir = os.path.abspath(work_dir)
    output_dir = os.path.abspath(output_dir)
    checks = []

    def add(name, ok, detail):
        checks.append({"name": name, "ok": bool(ok), "detail": detail})

    result_path = os.path.join(output_dir, "result.json")
    result = _j(result_path)
    add("result_json_parses", result is not None, "result.json exists and is JSON")
    if result is None:
        report = {"ok": False, "checks": checks, "generated_at": common.now_iso()}
        common.write_json(os.path.join(output_dir, "evidence", "validation_report.json"), report)
        return report

    manifest = _j(os.path.join(work_dir, "case_manifest.json"), {}) or {}
    required = ("identification", "delivery", "evidence", "slides", "plan",
                "template_inheritance", "open_items", "runtime", "output_contract")
    missing_keys = [k for k in required if k not in result]
    add("required_keys", not missing_keys, "missing=%s" % missing_keys)

    case_ok = (result.get("identification", {}).get("case_id") == manifest.get("case_id"))
    add("case_id_matches_input", case_ok,
        "result=%s manifest=%s" % (result.get("identification", {}).get("case_id"),
                                   manifest.get("case_id")))

    orig = manifest.get("output_contract") or {}
    got = result.get("output_contract")
    same = json.dumps(orig, ensure_ascii=False, sort_keys=True) == \
        json.dumps(got, ensure_ascii=False, sort_keys=True)
    add("output_contract_preserved_verbatim", same,
        "top-level keys: expected=%s got=%s" % (sorted(orig.keys()) if isinstance(orig, dict) else None,
                                                sorted(got.keys()) if isinstance(got, dict) else None))

    path_map = result.get("contract_path_map") or {}
    unfilled = (result.get("delivery", {}).get("audit", {}) or {}).get("unfilled_fields") or []
    leaves = _leaves(orig) if isinstance(orig, dict) else []
    unaccounted = [l for l in leaves if l not in path_map and l not in unfilled]
    add("contract_leaves_accounted", not unaccounted, "unaccounted=%s" % unaccounted)
    bad_targets = []
    for leaf, target in path_map.items():
        ok, val = _resolve(result, target)
        if not ok or val is None:
            bad_targets.append({"leaf": leaf, "target": target})
    add("contract_path_map_resolves", not bad_targets, "unresolved=%s" % bad_targets[:8])

    art_problems = []
    for a in (result.get("delivery", {}).get("artifacts") or []):
        rel = a.get("rel")
        p = os.path.join(output_dir, rel) if rel else None
        if not p or not os.path.isfile(p):
            art_problems.append({"rel": rel, "why": "missing file"})
            continue
        if a.get("sha256") and common.sha256_file(p) != a["sha256"]:
            art_problems.append({"rel": rel, "why": "sha256 mismatch"})
        if a.get("bytes") is not None and os.path.getsize(p) != a["bytes"]:
            art_problems.append({"rel": rel, "why": "byte size mismatch"})
    add("artifact_hashes_match_files", not art_problems, "problems=%s" % art_problems[:6])

    audit = (result.get("delivery", {}).get("audit") or {})
    status = (result.get("delivery", {}).get("status"))
    if status == "ok":
        ok_ok = (audit.get("fidelity_ok") is True and audit.get("format_status") == "pass"
                 and audit.get("render_status") == "rendered" and not unfilled)
        add("status_ok_is_earned", ok_ok,
            "ok requires fidelity=True format=pass render=rendered unfilled=[]; got %s" % audit)
    else:
        add("status_ok_is_earned", True, "status=%s (not claiming ok)" % status)
    if audit.get("fidelity_ok") is False:
        add("fidelity_failure_downgrades", status in (HARD_FAIL_STATUS, "partial"),
            "fidelity_ok=false but status=%s" % status)
    else:
        add("fidelity_failure_downgrades", True, "fidelity_ok=%s" % audit.get("fidelity_ok"))

    # CONTRACT v1.2: an unavailable renderer degrades to partial, never needs_human
    if audit.get("render_status") in ("skipped", "failed"):
        add("renderer_absent_is_partial", status != HARD_FAIL_STATUS,
            "render_status=%s -> status=%s (must not be needs_human)"
            % (audit.get("render_status"), status))
    else:
        add("renderer_absent_is_partial", True, "render_status=%s" % audit.get("render_status"))

    # needs_human only for: fidelity ok=false, critical format fail, or a missing auditor
    ev_now = result.get("evidence") or {}
    mods_present = True
    for key in ("fidelity", "format_audit"):
        entry = ev_now.get(key)
        if not (isinstance(entry, dict) and entry.get("rel")
                and os.path.isfile(os.path.join(output_dir, entry["rel"]))):
            mods_present = False
    if status == HARD_FAIL_STATUS:
        allowed = (audit.get("fidelity_ok") is False) or (audit.get("format_status") == "fail") \
            or (not mods_present)
        add("needs_human_only_for_hard_failures", allowed,
            "needs_human requires fidelity_ok=false | format fail | auditor missing; audit=%s"
            % audit)
    else:
        add("needs_human_only_for_hard_failures", True, "status=%s" % status)

    # rendered claim must be backed by a real PDF
    if audit.get("render_status") == "rendered":
        pdfs = [a for a in (result.get("delivery", {}).get("artifacts") or [])
                if a.get("kind") == "pdf"]
        good = False
        why = "no pdf artifact registered"
        for a in pdfs:
            p = os.path.join(output_dir, a["rel"])
            if os.path.isfile(p) and os.path.getsize(p) > 0:
                with open(p, "rb") as fh:
                    good = fh.read(5) == b"%PDF-"
                why = "%s signature=%s" % (a["rel"], good)
                if good:
                    break
        add("render_claim_backed_by_real_pdf", good, why)
        add("render_pages_registered", bool(result.get("evidence", {}).get("render")),
            "render evidence registered")
    else:
        add("render_claim_backed_by_real_pdf", True,
            "render_status=%s (no rendered claim)" % audit.get("render_status"))

    deck_rel = (result.get("delivery", {}).get("primary_output"))
    if deck_rel:
        deck_ok, deck_why = False, "missing"
        dp = os.path.join(output_dir, deck_rel)
        if os.path.isfile(dp):
            try:
                with zipfile.ZipFile(dp) as z:
                    n = len([x for x in z.namelist()
                             if re.match(r"^ppt/slides/slide\d+\.xml$", x)])
                    masters = [x for x in z.namelist()
                               if re.match(r"^ppt/slideMasters/slideMaster\d+\.xml$", x)]
                    deck_ok = n >= 1 and bool(masters)
                    deck_why = "slides=%d masters=%d" % (n, len(masters))
            except Exception as exc:  # noqa: BLE001
                deck_why = "%s: %s" % (type(exc).__name__, exc)
        add("deck_openable_with_inheritance", deck_ok, deck_why)
    else:
        # a null primary_output is honest for a needs_human delivery (e.g. unusable template)
        add("deck_openable_with_inheritance", status == HARD_FAIL_STATUS,
            "primary_output=null accepted only for needs_human (status=%s)" % status)

    # status-consistent limitations: an ok delivery must not claim nothing was produced
    lim = (result.get("open_items") or {}).get("limitations") or []
    conflicting = [l for l in lim if isinstance(l, str) and _STAGE_CLAIM_RX.search(l)]
    add("limitations_consistent_with_status", not (status == "ok" and conflicting),
        "status=%s conflicting=%s" % (status, conflicting[:3]))

    rt = result.get("runtime") or {}
    ps, fs = rt.get("prepare_s"), rt.get("finish_s")
    dup = (isinstance(ps, (int, float)) and isinstance(fs, (int, float)) and ps > 0 and ps == fs)
    add("runtime_stage_times_distinct", not dup,
        "prepare_s=%s finish_s=%s timing_source=%s" % (ps, fs, rt.get("timing_source")))

    # every evidence copy referenced must exist once run.py copied it
    ev_missing = []
    for key in ("sources", "fidelity", "format_audit", "render"):
        entry = (result.get("evidence") or {}).get(key)
        if isinstance(entry, dict) and entry.get("rel"):
            if not os.path.isfile(os.path.join(output_dir, entry["rel"])):
                ev_missing.append(entry["rel"])
    add("evidence_copies_present", not ev_missing, "missing=%s" % ev_missing)

    ok = all(c["ok"] for c in checks)
    report = {"ok": ok, "status": status, "checks": checks,
              "unfilled_fields": unfilled, "generated_at": common.now_iso(),
              "note": "ok=false means the delivery must not be presented as complete."}
    common.write_json(os.path.join(output_dir, "evidence", "validation_report.json"), report)
    return report


def main(argv=None):
    ap = argparse.ArgumentParser(description="finalize + validate output contract")
    ap.add_argument("--work", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--validate-only", action="store_true")
    args = ap.parse_args(argv)
    if args.validate_only:
        rep = validate(args.work, args.output)
    else:
        out = build(args.work, args.output)
        rep = validate(args.work, args.output)
        rep["build"] = {k: out[k] for k in ("status", "unfilled_fields")}
    print(json.dumps({"ok": rep.get("ok"), "status": rep.get("status"),
                      "checks": [{"name": c["name"], "ok": c["ok"]} for c in rep["checks"]]},
                     ensure_ascii=False, indent=2))
    return 0 if rep.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
