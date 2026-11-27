#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_report.py — real pptx -> PDF/PNG rendering evidence.

Frozen interface (CONTRACT §4.1)::

    render(work_dir, output_dir) -> dict
        writes <work>/render_report.json ; real artifacts go to <output>/preview/

Probe order (each attempt recorded in ``attempts[]``):

1. env ``SOFFICE_BIN``
2. ``soffice`` / ``libreoffice`` / ``libreoffice-kit`` on PATH
3. the DSH **LibreOffice Kit** (``node`` + ``cli.js``) — discovered at runtime via
   ``DSH_LOKIT_NODE``/``DSH_LOKIT_CLI``, ``DSH_LOKIT_ROOT``, ``dsoffice``, or a *bounded*
   scan of well-known DSH resource roots
4. ``office-renderer`` / ``unoconv`` / ``onlyoffice`` on PATH

No renderer -> ``status=skipped`` with the recorded attempts.  Never a fabricated preview:
every registered sha256/bytes/width/height is measured from a real file.
"""

from __future__ import annotations

import argparse
import glob as _glob
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import common  # noqa: E402

PDF_NAME = "slides.pdf"
KIT_TIMEOUT = 600
DPI = 96
MAX_PREVIEW_PAGES = 40


def _which(name):
    p = shutil.which(name)
    return p if p and os.path.isfile(p) else None


def _exe(path):
    if not path:
        return None
    p = os.path.abspath(path)
    return p if os.path.isfile(p) else None


def _bounded_glob(base, rel, max_depth, deadline):
    """Depth-limited glob (no unbounded ``**`` walk) with a wall-clock deadline."""
    hits = []
    for depth in range(0, max_depth + 1):
        if time.monotonic() > deadline:
            break
        pattern = os.path.join(base, *(["*"] * depth), rel)
        try:
            hits.extend(p for p in _glob.glob(pattern) if os.path.isfile(p))
        except Exception:
            pass
    return hits


_KIT_CLI_REL = os.path.join("@deepseek-ai", "libreoffice-kit*", "lib", "cli.js")
_NODE_RELS = (os.path.join("dependencies", "node", "bin", "node"),
              os.path.join("dependencies", "node", "bin", "node.exe"),
              "node", "node.exe")


def _kit_roots():
    """Candidate roots for an LibreOffice Kit, from env / PATH / this file only.

    No author-machine absolute path is ever embedded: every root is either an
    environment variable, a directory derived from a binary found on ``PATH``, or an
    ancestor of this file.  All scanning is depth-bounded.
    """
    roots = []  # (path, max_depth)

    def add(path, depth):
        if path:
            roots.append((path, depth))

    for var in ("DSH_LOKIT_ROOT", "DSH_RESOURCES", "DSH_OFFICE_ROOT", "BH_WORKSPACE"):
        add(os.environ.get(var), 6)
    add(os.environ.get("DSH_HOME"), 4)
    for var in ("APPDATA", "LOCALAPPDATA", "XDG_DATA_HOME", "HOME", "USERPROFILE"):
        base = os.environ.get(var)
        if base:
            add(os.path.join(base, "@deepseek-ai"), 3)
            add(os.path.join(base, ".local", "share"), 4)
    # reverse-derive from binaries that ship alongside the kit
    for name in ("dsoffice", "libreoffice-kit", "soffice", "libreoffice",
                 "office-renderer", "node", "node.exe"):
        exe = _which(name)
        if exe:
            exe = os.path.abspath(exe)
            add(os.path.dirname(exe), 5)
            parent = os.path.dirname(os.path.dirname(exe))
            add(parent, 5)
            grand = os.path.dirname(parent)
            add(grand, 5)
    # ancestors of this file (work/submission/scripts -> up to the project root)
    here = os.path.dirname(os.path.abspath(__file__))
    for depth in range(1, 7):
        add(os.path.join(here, *([".."] * depth)), 4)
    return roots


def _kit_candidates(deadline_s=20.0):
    """Bounded discovery of the DSH LibreOffice Kit (``node`` + ``cli.js``)."""
    deadline = time.monotonic() + deadline_s
    env_node = _exe(os.environ.get("DSH_LOKIT_NODE")) or _exe(os.environ.get("DSH_NODE_BIN"))
    env_cli = _exe(os.environ.get("DSH_LOKIT_CLI"))
    path_node = _which("node") or _which("node.exe")
    out = []
    if (env_node or path_node) and env_cli:
        out.append((env_node or path_node, env_cli))

    seen = set()
    for base, depth in _kit_roots():
        try:
            base = os.path.abspath(base)
        except Exception:
            continue
        if base in seen or not os.path.isdir(base):
            continue
        seen.add(base)
        hits = _bounded_glob(base, _KIT_CLI_REL, depth, deadline)
        if not hits:
            continue
        node = env_node or path_node
        if not node:
            for rel in _NODE_RELS:
                found = _bounded_glob(base, rel, depth, deadline)
                if found:
                    node = found[0]
                    break
        if not node:
            continue
        for hit in sorted(hits):
            out.append((node, hit))
        if out:
            break
    final, keys = [], set()
    for node, cli in out:
        k = (os.path.abspath(node), os.path.abspath(cli))
        if k in keys:
            continue
        keys.add(k)
        final.append({"source": "env" if os.environ.get("DSH_LOKIT_CLI") else "discovery",
                      "node": k[0], "cli": k[1]})
    return final


def _run(cmd, timeout=KIT_TIMEOUT, cwd=None):
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout, cwd=cwd)
        return proc.returncode, proc.stdout.decode("utf-8", "replace"), \
            proc.stderr.decode("utf-8", "replace")
    except subprocess.TimeoutExpired:
        return None, "", "timeout after %ss" % timeout
    except Exception as exc:  # noqa: BLE001
        return None, "", "%s: %s" % (type(exc).__name__, exc)


def _json_from_stdout(text):
    for line in reversed((text or "").splitlines()):
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except ValueError:
                continue
    return None


def _pdf_page_count(data: bytes):
    if not data.startswith(b"%PDF-"):
        return None
    n = len(re.findall(rb"/Type\s*/Page[^s]", data))
    if n:
        return n
    counts = [int(x) for x in re.findall(rb"/Count\s+(\d+)", data)]
    return max(counts) if counts else None


def _pdfinfo_pages(pdf_path):
    tool = _which("pdfinfo")
    if not tool:
        return None
    rc, out, _ = _run([tool, pdf_path], timeout=60)
    if rc != 0:
        return None
    m = re.search(r"^Pages:[ \t]+(\d+)", out, re.M)
    return int(m.group(1)) if m else None


def _pdffonts(pdf_path):
    tool = _which("pdffonts")
    if not tool:
        return None
    rc, out, _ = _run([tool, pdf_path], timeout=60)
    if rc != 0:
        return None
    embedded, subs = [], []
    lines = out.splitlines()
    for line in lines[2:]:
        if not line.strip():
            continue
        cols = re.split(r"\s{2,}", line.strip())
        if len(cols) < 5:
            continue
        name, ftype = cols[0], cols[1]
        emb = "yes" in cols[-4].lower() if len(cols) >= 4 else False
        embedded.append({"name": name, "type": ftype, "embedded": bool(emb)})
    return {"tool": "pdffonts", "embedded": embedded, "substitutions": subs,
            "raw_lines": lines[:12]}


def _pptx_slide_count(deck_path):
    try:
        with zipfile.ZipFile(deck_path) as z:
            return len([n for n in z.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)])
    except Exception:
        return None


def _png_size(path):
    return common.natural_size(path)


def _register_file(path, rel):
    size = os.path.getsize(path)
    return {"rel": rel, "sha256": common.sha256_file(path), "bytes": size}


def _find_deck(output_dir, work_dir):
    plan = common.read_json(os.path.join(work_dir, "slides_plan.json"), default={}) or {}
    name = plan.get("deck_name") or common.DECK_DEFAULT_NAME
    cand = os.path.join(output_dir, name)
    if os.path.isfile(cand):
        return cand
    found = []
    for root, dirs, files in os.walk(output_dir):
        dirs[:] = [d for d in dirs if d not in ("preview", "__pycache__")]
        for f in sorted(files):
            if f.lower().endswith((".pptx", ".pptm")) and not f.startswith("~$"):
                found.append(os.path.join(root, f))
    for f in found:
        if os.path.basename(f) == name:
            return f
    return found[0] if found else None


def _try_soffice(binary, deck_path, preview_dir, attempts):
    """soffice-style CLI: --headless --convert-to pdf --outdir DIR DECK."""
    rf = os.path.join(preview_dir, "." + os.path.basename(binary) + "_profile")
    cmd = [binary, "--headless", "--norestore", "--nolockcheck",
           "-env:UserInstallation=file:///" + rf.replace("\\", "/"),
           "--convert-to", "pdf", "--outdir", preview_dir, deck_path]
    produced = os.path.join(preview_dir, os.path.splitext(os.path.basename(deck_path))[0] + ".pdf")
    if os.path.exists(produced):
        # same re-entrancy guard as the kit path: a stale target must not downgrade the run
        try:
            os.remove(produced)
        except OSError:
            shutil.rmtree(preview_dir, ignore_errors=True)
            os.makedirs(preview_dir, exist_ok=True)
    rc, out, err = _run(cmd)
    detail = "rc=%s out=%s err=%s" % (rc, out.strip()[:200], err.strip()[:200])
    attempts.append({"renderer": os.path.basename(binary), "status": "ok" if rc == 0 else "failed",
                     "detail": detail})
    if rc == 0 and os.path.isfile(produced) and os.path.getsize(produced) > 0:
        return {"kind": "soffice", "name": os.path.basename(binary), "path": binary,
                "pdf": produced, "info": {"stdout": out[:400]}}
    return None


def _try_kit(cand, deck_path, preview_dir, attempts):
    node, cli = cand["node"], cand["cli"]
    raster_dir = preview_dir + ".raster"
    if os.path.exists(raster_dir):
        shutil.rmtree(raster_dir, ignore_errors=True)
    rc, out, err = _run([node, cli, "capabilities", "--json"], timeout=120)
    caps = _json_from_stdout(out) or {}
    attempts.append({"renderer": "libreoffice-kit:capabilities",
                     "status": "ok" if rc == 0 else "failed",
                     "detail": "cli=%s version=%s backend=%s" %
                               (cli, (caps.get("runtime") or {}).get("version"),
                                (caps.get("runtime") or {}).get("backend"))})
    rc_r, out_r, err_r = _run([node, cli, "render", "--input", deck_path,
                               "--output-dir", raster_dir, "--pages", "all",
                               "--dpi", str(DPI)])
    rinfo = _json_from_stdout(out_r) or {}
    attempts.append({"renderer": "libreoffice-kit:render",
                     "status": "ok" if rc_r == 0 else "failed",
                     "detail": "rc=%s pages=%s err=%s" % (rc_r, rinfo.get("pageCount"),
                                                          (err_r or "").strip()[:200])})
    pdf_out = os.path.join(preview_dir, PDF_NAME)
    # the renderer refuses to overwrite an existing target (EEXIST -> "skipped", which would
    # silently downgrade delivery to partial on a re-run in the same directory)
    if os.path.exists(pdf_out):
        try:
            os.remove(pdf_out)
        except OSError:
            shutil.rmtree(preview_dir, ignore_errors=True)
            os.makedirs(preview_dir, exist_ok=True)
    rc_c, out_c, err_c = _run([node, cli, "convert", "--input", deck_path, "--output", pdf_out])
    cinfo = _json_from_stdout(out_c) or {}
    attempts.append({"renderer": "libreoffice-kit:convert",
                     "status": "ok" if rc_c == 0 else "failed",
                     "detail": "rc=%s missingFonts=%s err=%s"
                               % (rc_c, cinfo.get("missingFonts"), (err_c or "").strip()[:200])})
    if rc_c != 0 or not os.path.isfile(pdf_out) or os.path.getsize(pdf_out) == 0:
        return None
    return {"kind": "kit", "name": "office-renderer", "path": cli, "node": node,
            "pdf": pdf_out, "raster_dir": raster_dir, "render_info": rinfo,
            "convert_info": cinfo, "capabilities": caps}


def _try_cli_stub(binary, deck_path, preview_dir, attempts):
    """office-renderer / unoconv / onlyoffice best-effort (recorded either way)."""
    base = os.path.basename(binary).lower()
    if "unoconv" in base:
        cmd = [binary, "-f", "pdf", "-o", preview_dir, deck_path]
    else:
        cmd = [binary, "--input", deck_path, "--output", os.path.join(preview_dir, PDF_NAME)]
    rc, out, err = _run(cmd)
    attempts.append({"renderer": base, "status": "ok" if rc == 0 else "failed",
                     "detail": "rc=%s err=%s" % (rc, (err or "").strip()[:200])})
    if rc == 0:
        cand_pdf = os.path.join(preview_dir, PDF_NAME)
        if os.path.isfile(cand_pdf) and os.path.getsize(cand_pdf) > 0:
            return {"kind": base, "name": base, "path": binary, "pdf": cand_pdf, "info": {}}
    return None


def _rasterize_with_tool(pdf_path, preview_dir):
    """Rasterize via pdftoppm/pdftocairo/mutool if one is available."""
    tool = _which("pdftoppm") or _which("pdftocairo")
    if tool:
        stem = os.path.join(preview_dir, "page")
        rc, out, err = _run([tool, "-png", "-r", str(DPI), pdf_path, stem], timeout=300)
        if rc == 0:
            pages = []
            for p in sorted(_glob.glob(stem + "*.png")):
                sz = _png_size(p)
                rel = "preview/" + os.path.basename(p)
                if sz:
                    pages.append({"n": len(pages) + 1, "rel": rel,
                                  "sha256": common.sha256_file(p),
                                  "bytes": os.path.getsize(p),
                                  "width": sz[0], "height": sz[1]})
            return pages
    mutool = _which("mutool")
    if mutool:
        outpat = os.path.join(preview_dir, "page-%03d.png")
        rc, out, err = _run([mutool, "draw", "-r", str(DPI), "-o", outpat, pdf_path], timeout=300)
        if rc == 0:
            pages = []
            for p in sorted(_glob.glob(os.path.join(preview_dir, "page-*.png"))):
                sz = _png_size(p)
                if sz:
                    pages.append({"n": len(pages) + 1, "rel": "preview/" + os.path.basename(p),
                                  "sha256": common.sha256_file(p),
                                  "bytes": os.path.getsize(p),
                                  "width": sz[0], "height": sz[1]})
            return pages
    return []


def _ingest_kit_pages(raster_dir, preview_dir):
    """Move kit-rendered pages into <output>/preview and measure them for real."""
    manifest = common.read_json(os.path.join(raster_dir, "manifest.json"), default={}) or {}
    images = manifest.get("images") or []
    pages = []
    for i, img in enumerate(images, 1):
        src = img.get("path")
        if not src or not os.path.isfile(src):
            candidate = os.path.join(raster_dir, os.path.basename(src or ""))
            src = candidate if os.path.isfile(candidate) else None
        if not src:
            continue
        dst = os.path.join(preview_dir, "page-%03d.png" % i)
        shutil.copyfile(src, dst)
        sz = _png_size(dst)
        if not sz:
            continue
        pages.append({"n": i, "rel": "preview/" + os.path.basename(dst),
                      "sha256": common.sha256_file(dst), "bytes": os.path.getsize(dst),
                      "width": sz[0], "height": sz[1]})
        if len(pages) >= MAX_PREVIEW_PAGES:
            break
    return pages, manifest


def render(work_dir: str, output_dir: str) -> dict:
    """Render the deck to a real PDF (and PNG pages when a rasterizer exists)."""
    work_dir = os.path.abspath(work_dir)
    output_dir = os.path.abspath(output_dir)
    preview_dir = os.path.join(output_dir, "preview")
    common.ensure_dir(preview_dir)
    attempts = []
    report = {
        "status": "skipped",
        "renderer": {"name": None, "path": None, "version": None},
        "artifacts": [],
        "pages": [],
        "fonts": {"tool": "none", "embedded": [], "substitutions": []},
        "consistency": {"pptx_slide_count": None, "pdf_page_count": None, "ok": False},
        "attempts": attempts,
        "generated_at": common.now_iso(),
        "note": "Only real, non-empty files are registered; no renderer means skipped, never a fake preview.",
    }

    deck_path = _find_deck(output_dir, work_dir)
    report["deck"] = deck_path
    report["consistency"]["pptx_slide_count"] = _pptx_slide_count(deck_path) if deck_path else None
    if not deck_path:
        report["reason"] = "no .pptx found under output_dir"
        attempts.append({"renderer": "deck", "status": "failed", "detail": report["reason"]})
        common.write_json(os.path.join(work_dir, "render_report.json"), report)
        return report

    chosen = None
    soffice = _exe(os.environ.get("SOFFICE_BIN")) or _which("soffice") or _which("soffice.exe") \
        or _which("libreoffice") or _which("libreoffice-kit")
    if soffice:
        chosen = _try_soffice(soffice, deck_path, preview_dir, attempts)

    if chosen is None:
        for cand in _kit_candidates():
            attempts.append({"renderer": "libreoffice-kit:discover", "status": "ok",
                             "detail": "node=%s cli=%s source=%s"
                                       % (cand["node"], cand["cli"], cand["source"])})
            chosen = _try_kit(cand, deck_path, preview_dir, attempts)
            if chosen:
                break

    if chosen is None:
        for name in ("office-renderer", "unoconv", "onlyoffice"):
            b = _which(name)
            if b:
                chosen = _try_cli_stub(b, deck_path, preview_dir, attempts)
                if chosen:
                    break

    if chosen is None:
        report["status"] = "skipped"
        report["reason"] = "no pptx->pdf renderer found (see attempts)"
        common.write_json(os.path.join(work_dir, "render_report.json"), report)
        return report

    pdf_path = chosen["pdf"]
    try:
        with open(pdf_path, "rb") as fh:
            head = fh.read(8)
    except Exception:
        head = b""
    if not head.startswith(b"%PDF-") or os.path.getsize(pdf_path) == 0:
        report["status"] = "failed"
        report["reason"] = "renderer produced a non-PDF artifact"
        attempts.append({"renderer": "verify:%s" % chosen["kind"], "status": "failed",
                         "detail": "missing %%PDF- signature or empty file"})
        common.write_json(os.path.join(work_dir, "render_report.json"), report)
        return report

    final_pdf = os.path.join(preview_dir, PDF_NAME)
    if os.path.abspath(pdf_path) != os.path.abspath(final_pdf):
        shutil.copyfile(pdf_path, final_pdf)
    with open(final_pdf, "rb") as fh:
        pdf_bytes = fh.read()
    report["renderer"] = {"name": chosen["name"], "path": chosen["path"],
                          "version": ((chosen.get("capabilities") or {}).get("runtime") or {}).get("version")
                          or ((chosen.get("convert_info") or {}).get("version"))}
    report["artifacts"].append(_register_file(final_pdf, "preview/" + PDF_NAME))

    # pages: prefer the kit manifest (measured), else rasterize the real PDF
    pages, kit_manifest = ([], {})
    if chosen.get("raster_dir"):
        pages, kit_manifest = _ingest_kit_pages(chosen["raster_dir"], preview_dir)
    if not pages:
        pages = _rasterize_with_tool(final_pdf, preview_dir)
    report["pages"] = pages
    if kit_manifest:
        report["raster_manifest"] = {
            "dpi": kit_manifest.get("dpi"), "pageCount": kit_manifest.get("pageCount"),
            "sourceSha256": kit_manifest.get("sourceSha256"),
            "missingFonts": kit_manifest.get("missingFonts"),
        }

    pdf_pages = _pdf_page_count(pdf_bytes) or _pdfinfo_pages(final_pdf)
    report["consistency"]["pdf_page_count"] = pdf_pages
    pptx_n = report["consistency"]["pptx_slide_count"]
    equal = pdf_pages is not None and pptx_n is not None and pdf_pages == pptx_n
    if pages:
        equal = equal and len(pages) == pdf_pages
    report["consistency"]["ok"] = bool(equal)

    fonts = _pdffonts(final_pdf)
    missing = []
    if kit_manifest.get("missingFonts"):
        missing = list(kit_manifest["missingFonts"])
    if fonts:
        report["fonts"] = fonts
    if missing:
        subs = [{"requested": m, "used": None, "evidence": "libreoffice-kit render json missingFonts"}
                for m in missing]
        report["fonts"] = {"tool": report["fonts"].get("tool", "none"),
                           "embedded": report["fonts"].get("embedded", []),
                           "substitutions": report["fonts"].get("substitutions", []) + subs,
                           "missing_fonts": missing}
    if (chosen.get("convert_info") or {}).get("missingFonts"):
        report["fonts"]["missing_fonts"] = chosen["convert_info"]["missingFonts"]

    report["status"] = "rendered"
    report["forge_check"] = {
        "pdf_signature": True, "pdf_bytes": os.path.getsize(final_pdf),
        "pages_measured": len(pages), "pptx_slide_count": pptx_n, "pdf_page_count": pdf_pages,
    }
    common.write_json(os.path.join(work_dir, "render_report.json"), report)
    return report


def main(argv=None):
    ap = argparse.ArgumentParser(description="real pptx render evidence")
    ap.add_argument("--work", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args(argv)
    rep = render(args.work, args.output)
    print(json.dumps({"status": rep.get("status"),
                      "renderer": rep.get("renderer"), "pages": len(rep.get("pages") or []),
                      "artifacts": rep.get("artifacts"), "consistency": rep.get("consistency"),
                      "reason": rep.get("reason")}, ensure_ascii=False, indent=2))
    return 0 if rep.get("status") in ("rendered", "skipped") else 1


if __name__ == "__main__":
    raise SystemExit(main())
