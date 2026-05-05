#!/usr/bin/env python3
"""
converter_ui.py — UI web local para el conversor a Markdown
Ejecutar: python converter_ui.py
Luego abrir: http://localhost:3200
"""

import os
import sys
import json
import time
import shutil
import threading
import webbrowser
import tempfile
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string

# Asegurarse que convert_to_md está en el path
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from convert_to_md import convert_docx, convert_pdf, convert_html, convert_xlsx, convert_csv, convert_pptx, convert_eml, convert_msg, convert_with_markitdown, MARKITDOWN_EXTENSIONS, convert_image, IMAGE_EXTENSIONS

app = Flask(__name__)
OUTPUT_DIR = SCRIPT_DIR / "md_output"
OUTPUT_DIR.mkdir(exist_ok=True)

DEFAULT_WATCH_DIR = r"D:\mdconverter\correos"
WATCH_DIR = None
WATCH_THREAD = None
WATCH_ACTIVE = False


def _start_watcher(folder_path: str) -> bool:
    """Inicia el watcher en background para una carpeta dada."""
    global WATCH_DIR, WATCH_THREAD, WATCH_ACTIVE
    path = Path(folder_path)
    if not path.exists():
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception:
            return False
    WATCH_ACTIVE = False  # detener watcher anterior si existe
    time.sleep(0.2)
    WATCH_DIR = path
    WATCH_ACTIVE = True

    def _run():
        seen = set(WATCH_DIR.glob("*"))
        while WATCH_ACTIVE:
            time.sleep(1.5)
            current = set(WATCH_DIR.glob("*"))
            new_files = current - seen
            seen = current
            for f in new_files:
                if f.suffix.lower() in SUPPORTED:
                    do_convert(f, OUTPUT_DIR)

    WATCH_THREAD = threading.Thread(target=_run, daemon=True)
    WATCH_THREAD.start()
    return True

# ─── Conversión ───────────────────────────────────────────────────────────────

SUPPORTED = {".docx", ".pdf", ".pptx", ".html", ".htm", ".xlsx", ".csv", ".eml", ".msg"} | MARKITDOWN_EXTENSIONS | IMAGE_EXTENSIONS

def do_convert(src: Path, out_dir: Path) -> list[dict]:
    ext = src.suffix.lower()
    results = []
    try:
        if ext in (".eml", ".msg"):
            fn = convert_eml if ext == ".eml" else convert_msg
            pairs = fn(src)
            for fname, content in pairs:
                dest = out_dir / fname
                dest.write_text(content, encoding="utf-8")
                results.append({"name": fname, "path": str(dest), "ok": True})
        else:
            if ext == ".docx":             content = convert_docx(src)
            elif ext == ".pptx":           content = convert_pptx(src)
            elif ext == ".pdf":            content = convert_pdf(src)
            elif ext in (".html", ".htm"): content = convert_html(str(src))
            elif ext == ".xlsx":           content = convert_xlsx(src)
            elif ext == ".csv":            content = convert_csv(src)
            elif ext in IMAGE_EXTENSIONS:  content = convert_image(src)
            elif ext in MARKITDOWN_EXTENSIONS: content = convert_with_markitdown(src)
            else:
                return [{"name": src.name, "ok": False, "error": f"Formato no soportado: {ext}"}]
            dest = out_dir / f"{src.stem}.md"
            dest.write_text(content, encoding="utf-8")
            results.append({"name": dest.name, "path": str(dest), "ok": True})
    except Exception as e:
        results.append({"name": src.name, "ok": False, "error": str(e)})
    return results


# ─── Rutas API ────────────────────────────────────────────────────────────────

@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No se recibieron archivos"}), 400

    all_results = []
    for f in files:
        suffix = Path(f.filename).suffix.lower()
        if suffix not in SUPPORTED:
            all_results.append({"name": f.filename, "ok": False,
                                 "error": f"Formato no soportado: {suffix}"})
            continue
        tmp = Path(tempfile.mktemp(suffix=suffix))
        f.save(tmp)
        results = do_convert(tmp, OUTPUT_DIR)
        all_results.extend(results)
        tmp.unlink(missing_ok=True)

    return jsonify(all_results)


@app.route("/preview/<path:filename>")
def preview(filename):
    fp = OUTPUT_DIR / filename
    if not fp.exists():
        return jsonify({"error": "Archivo no encontrado"}), 404
    return jsonify({"content": fp.read_text(encoding="utf-8")})


@app.route("/download/<path:filename>")
def download(filename):
    fp = OUTPUT_DIR / filename
    if not fp.exists():
        return "Archivo no encontrado", 404
    return send_file(fp, as_attachment=True)


@app.route("/watch/start", methods=["POST"])
def watch_start():
    global WATCH_ACTIVE
    data = request.json or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Carpeta no válida"}), 400
    WATCH_ACTIVE = False  # stop existing watcher
    import time; time.sleep(0.2)
    ok = _start_watcher(folder)
    if not ok:
        return jsonify({"error": f"No se pudo crear/acceder a: {folder}"}), 400
    return jsonify({"ok": True, "folder": str(WATCH_DIR)})


@app.route("/watch/stop", methods=["POST"])
def watch_stop():
    global WATCH_ACTIVE
    WATCH_ACTIVE = False
    return jsonify({"ok": True})


@app.route("/files")
def list_files():
    files = sorted(OUTPUT_DIR.glob("*.md"), key=lambda f: f.stat().st_mtime, reverse=True)
    return jsonify([{
        "name": f.name,
        "size_kb": round(f.stat().st_size / 1024, 1),
        "mtime": f.stat().st_mtime
    } for f in files[:50]])


@app.route("/convert-clipboard", methods=["POST"])
def convert_clipboard():
    from converters.html import _html_to_md_with_tables
    data = request.json or {}
    kind = data.get("type", "text")   # "html" | "text"
    content = data.get("content", "")

    if not content.strip():
        return jsonify({"error": "El portapapeles está vacío"}), 400

    stem = f"clipboard_{int(time.time())}"
    try:
        md = _html_to_md_with_tables(content) if kind == "html" else content
        fname = f"{stem}.md"
        dest = OUTPUT_DIR / fname
        dest.write_text(md, encoding="utf-8")
        return jsonify([{"name": fname, "path": str(dest), "ok": True}])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─── UI ───────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MD Converter</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@400;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
<style>
  :root {
    --bg:      #282a36;
    --surface: #21222c;
    --border:  #44475a;
    --accent:  #bd93f9;
    --accent2: #8be9fd;
    --warn:    #ff5555;
    --text:    #f8f8f2;
    --muted:   #6272a4;
    --green:   #50fa7b;
    --pink:    #ff79c6;
    --orange:  #ffb86c;
    --mono:    'DM Mono', monospace;
    --display: 'Syne', sans-serif;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--mono);
    font-size: 13px;
    min-height: 100vh;
    display: grid;
    grid-template-rows: auto 1fr;
    overflow-x: hidden;
  }

  /* noise overlay */
  body::before {
    content: '';
    position: fixed; inset: 0;
    background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.04'/%3E%3C/svg%3E");
    pointer-events: none; z-index: 0;
  }

  header {
    padding: 20px 32px;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    gap: 16px;
    position: relative; z-index: 1;
  }

  .logo {
    font-family: var(--display);
    font-size: 22px;
    font-weight: 800;
    letter-spacing: -0.5px;
    color: var(--accent);
  }

  .logo span { color: var(--text); }

  .badge {
    background: var(--surface);
    border: 1px solid var(--border);
    padding: 2px 8px;
    border-radius: 3px;
    font-size: 11px;
    color: var(--muted);
    letter-spacing: 0.05em;
  }

  main {
    display: grid;
    grid-template-columns: 1fr 280px;
    gap: 0;
    position: relative; z-index: 1;
    height: calc(100vh - 61px);
  }

  /* ── LEFT PANEL ── */
  .left {
    padding: 32px;
    display: flex;
    flex-direction: column;
    gap: 24px;
    overflow-y: auto;
    border-right: 1px solid var(--border);
  }

  /* Drop zone */
  .dropzone {
    border: 2px dashed var(--border);
    border-radius: 8px;
    padding: 52px 32px;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s ease;
    position: relative;
    background: var(--surface);
  }

  .dropzone::before {
    content: '';
    position: absolute; inset: 0;
    border-radius: 6px;
    background: radial-gradient(ellipse at 50% 0%, rgba(189,147,249,0.04) 0%, transparent 70%);
    pointer-events: none;
  }

  .dropzone.drag-over {
    border-color: var(--accent);
    background: rgba(189,147,249,0.05);
    transform: scale(1.005);
  }

  .dropzone.drag-over::before {
    background: radial-gradient(ellipse at 50% 0%, rgba(189,147,249,0.12) 0%, transparent 70%);
  }

  .drop-icon {
    font-size: 40px;
    margin-bottom: 16px;
    display: block;
    transition: transform 0.2s;
  }

  .dropzone.drag-over .drop-icon { transform: scale(1.15) translateY(-4px); }

  .drop-title {
    font-family: var(--display);
    font-size: 18px;
    font-weight: 700;
    color: var(--text);
    margin-bottom: 8px;
  }

  .drop-sub {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.6;
  }

  .drop-sub strong { color: var(--accent); }

  .file-types {
    display: flex;
    justify-content: center;
    gap: 6px;
    margin-top: 16px;
    flex-wrap: wrap;
  }

  .ft {
    padding: 3px 8px;
    border-radius: 3px;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.05em;
  }

  .ft-docx { background: rgba(0,102,255,0.15); color: #4d9fff; border: 1px solid rgba(0,102,255,0.25); }
  .ft-pdf  { background: rgba(255,77,109,0.15); color: #ff6b8a; border: 1px solid rgba(255,77,109,0.25); }
  .ft-xlsx { background: rgba(0,229,160,0.12); color: var(--accent); border: 1px solid rgba(0,229,160,0.2); }
  .ft-html { background: rgba(255,165,0,0.12); color: #ffb347; border: 1px solid rgba(255,165,0,0.2); }
  .ft-eml  { background: rgba(160,100,255,0.15); color: #c084fc; border: 1px solid rgba(160,100,255,0.25); }
  .ft-csv  { background: rgba(100,200,255,0.12); color: #67e3ff; border: 1px solid rgba(100,200,255,0.2); }
  .ft-img  { background: rgba(255,200,50,0.12); color: #fcd34d; border: 1px solid rgba(255,200,50,0.25); }
  .ft-epub { background: rgba(255,184,108,0.10); color: var(--orange);  border: 1px solid rgba(255,184,108,0.2); }
  .ft-json { background: rgba(80,250,123,0.10);  color: var(--green);   border: 1px solid rgba(80,250,123,0.18); }
  .ft-xml  { background: rgba(139,233,253,0.08); color: var(--accent2); border: 1px solid rgba(139,233,253,0.18); }
  .ft-zip  { background: rgba(189,147,249,0.10); color: var(--accent);  border: 1px solid rgba(189,147,249,0.18); }

  #file-input { display: none; }

  /* Results */
  .section-title {
    font-family: var(--display);
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--muted);
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .section-title::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
  }

  .results-list {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .result-item {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
    display: flex;
    align-items: center;
    gap: 12px;
    animation: slideIn 0.25s ease;
  }

  @keyframes slideIn {
    from { opacity: 0; transform: translateY(8px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  .result-item.ok { border-left: 3px solid var(--accent); }
  .result-item.err { border-left: 3px solid var(--warn); }
  .result-item.converting { border-left: 3px solid var(--accent2); opacity: 0.7; }

  .result-icon { font-size: 16px; flex-shrink: 0; }

  .result-info { flex: 1; min-width: 0; }

  .result-name {
    color: var(--text);
    font-weight: 500;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    font-size: 13px;
  }

  .result-meta {
    color: var(--muted);
    font-size: 11px;
    margin-top: 2px;
  }

  .result-meta.err-msg { color: var(--warn); }

  .btn-dl {
    background: rgba(189,147,249,0.1);
    border: 1px solid rgba(189,147,249,0.25);
    color: var(--accent);
    padding: 5px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    text-decoration: none;
    flex-shrink: 0;
    transition: all 0.15s;
    white-space: nowrap;
  }
  .btn-dl:hover { background: rgba(189,147,249,0.2); border-color: var(--accent); }

  .btn-copy {
    background: rgba(255,255,255,0.04);
    border: 1px solid var(--border);
    color: var(--muted);
    padding: 5px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    flex-shrink: 0;
    transition: all 0.15s;
    white-space: nowrap;
  }
  .btn-copy:hover { background: rgba(255,255,255,0.08); color: var(--text); border-color: var(--muted); }
  .btn-copy.copied { background: rgba(189,147,249,0.1); border-color: rgba(189,147,249,0.25); color: var(--accent); }

  /* Spinner */
  .spinner {
    width: 14px; height: 14px;
    border: 2px solid var(--border);
    border-top-color: var(--accent2);
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
    flex-shrink: 0;
  }

  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── RIGHT PANEL ── */
  .right {
    padding: 24px;
    display: flex;
    flex-direction: column;
    gap: 20px;
    overflow-y: auto;
    background: var(--surface);
  }

  /* Watch folder */
  .watch-card {
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 16px;
  }

  .watch-title {
    font-family: var(--display);
    font-size: 13px;
    font-weight: 700;
    color: var(--text);
    margin-bottom: 4px;
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .watch-desc {
    color: var(--muted);
    font-size: 11px;
    line-height: 1.5;
    margin-bottom: 12px;
  }

  .watch-desc strong { color: var(--accent); }

  .input-row {
    display: flex;
    gap: 6px;
    margin-bottom: 8px;
  }

  .input-folder {
    flex: 1;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--text);
    padding: 7px 10px;
    border-radius: 4px;
    font-family: var(--mono);
    font-size: 11px;
    outline: none;
    transition: border-color 0.15s;
  }

  .input-folder:focus { border-color: var(--accent); }

  .btn-primary {
    background: var(--accent);
    color: var(--bg);
    border: none;
    padding: 7px 14px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    transition: opacity 0.15s;
    white-space: nowrap;
  }

  .btn-primary:hover { opacity: 0.85; }

  .btn-danger {
    background: var(--warn);
    color: #fff;
    border: none;
    padding: 7px 14px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    transition: opacity 0.15s;
  }

  .btn-danger:hover { opacity: 0.85; }

  .watch-status {
    font-size: 11px;
    padding: 6px 10px;
    border-radius: 4px;
    display: none;
  }

  .watch-status.active {
    display: block;
    background: rgba(189,147,249,0.1);
    border: 1px solid rgba(189,147,249,0.2);
    color: var(--accent);
  }

  .watch-status.inactive {
    display: block;
    background: rgba(255,85,85,0.1);
    border: 1px solid rgba(255,85,85,0.2);
    color: var(--warn);
  }

  .pulse {
    display: inline-block;
    width: 7px; height: 7px;
    background: var(--accent);
    border-radius: 50%;
    animation: pulse 1.5s ease-in-out infinite;
    margin-right: 6px;
  }

  @keyframes pulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%       { opacity: 0.4; transform: scale(0.7); }
  }

  /* File history */
  .history-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .hist-item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 7px 10px;
    border-radius: 4px;
    background: var(--bg);
    border: 1px solid var(--border);
    transition: border-color 0.15s;
  }

  .hist-item:hover { border-color: var(--muted); }

  .hist-name {
    flex: 1;
    font-size: 11px;
    color: var(--text);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .hist-size {
    font-size: 10px;
    color: var(--muted);
    flex-shrink: 0;
  }

  .hist-dl {
    font-size: 14px;
    color: var(--muted);
    text-decoration: none;
    flex-shrink: 0;
    transition: color 0.15s;
  }

  .hist-dl:hover { color: var(--accent); }

  .empty-state {
    color: var(--muted);
    font-size: 11px;
    text-align: center;
    padding: 16px;
  }

  .btn-open {
    flex: 1;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--muted);
    padding: 6px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    transition: all 0.15s;
    text-align: left;
  }

  .btn-open:hover {
    border-color: var(--accent);
    color: var(--accent);
  }


  /* Clipboard paste button */
  .btn-clipboard {
    width: 100%;
    background: rgba(139,233,253,0.05);
    border: 1px dashed rgba(139,233,253,0.28);
    color: var(--accent2);
    padding: 10px 16px;
    border-radius: 6px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 12px;
    font-weight: 500;
    transition: all 0.15s;
    text-align: center;
  }
  .btn-clipboard:hover:not(:disabled) {
    background: rgba(139,233,253,0.11);
    border-color: var(--accent2);
  }
  .btn-clipboard:disabled { opacity: 0.45; cursor: not-allowed; }
  .btn-clipboard.pasting {
    border-color: var(--accent2);
    background: rgba(139,233,253,0.09);
    animation: clipPulse 0.9s ease-in-out infinite;
  }
  @keyframes clipPulse {
    0%, 100% { opacity: 1; }
    50%       { opacity: 0.55; }
  }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }

  /* ── Filter ── */
  .filter-wrap { position: relative; }
  .filter-input {
    width: 100%;
    background: var(--bg);
    border: 1px solid var(--border);
    color: var(--text);
    padding: 7px 10px 7px 30px;
    border-radius: 4px;
    font-family: var(--mono);
    font-size: 11px;
    outline: none;
    transition: border-color 0.15s;
  }
  .filter-input:focus { border-color: var(--accent); }
  .filter-input::placeholder { color: var(--muted); }
  .filter-icon {
    position: absolute;
    left: 9px; top: 50%;
    transform: translateY(-50%);
    font-size: 12px;
    color: var(--muted);
    pointer-events: none;
  }

  /* ── Copy button ── */
  .hist-copy {
    font-size: 13px;
    background: none;
    border: none;
    color: var(--muted);
    cursor: pointer;
    flex-shrink: 0;
    transition: color 0.15s;
    padding: 0 2px;
    line-height: 1;
  }
  .hist-copy:hover { color: var(--accent); }
  .hist-copy.copied { color: var(--accent); }

  /* ── Preview modal ── */
  .preview-overlay {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,0.75);
    z-index: 100;
    backdrop-filter: blur(4px);
    align-items: flex-start;
    justify-content: center;
    padding: 32px 24px;
  }
  .preview-overlay.open { display: flex; }

  .preview-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    width: 100%;
    max-width: 820px;
    max-height: calc(100vh - 64px);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    animation: slideIn 0.2s ease;
  }

  .preview-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 14px 20px;
    border-bottom: 1px solid var(--border);
    flex-shrink: 0;
  }

  .preview-title {
    flex: 1;
    font-size: 11px;
    color: var(--muted);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .preview-btn {
    background: none;
    border: 1px solid var(--border);
    color: var(--muted);
    padding: 4px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    transition: all 0.15s;
    flex-shrink: 0;
  }
  .preview-btn:hover { border-color: var(--accent); color: var(--accent); }
  .preview-btn.copied { border-color: var(--accent); color: var(--accent); }

  .preview-close {
    background: none;
    border: none;
    color: var(--muted);
    font-size: 18px;
    cursor: pointer;
    line-height: 1;
    padding: 0 2px;
    transition: color 0.15s;
    flex-shrink: 0;
  }
  .preview-close:hover { color: var(--warn); }

  .preview-content {
    padding: 28px 32px;
    overflow-y: auto;
    flex: 1;
    line-height: 1.75;
    font-size: 13px;
    color: var(--text);
  }

  .preview-fm {
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 14px;
    margin-bottom: 20px;
    font-size: 11px;
    color: var(--muted);
    white-space: pre-wrap;
    line-height: 1.6;
  }

  /* Markdown rendered */
  .preview-content h1 { font-family: var(--display); font-size: 20px; font-weight: 700; margin: 0 0 16px; color: var(--accent); }
  .preview-content h2 { font-size: 14px; font-weight: 700; margin: 24px 0 10px; color: var(--text); border-bottom: 1px solid var(--border); padding-bottom: 6px; text-transform: uppercase; letter-spacing: 0.05em; }
  .preview-content h3 { font-size: 13px; font-weight: 600; margin: 16px 0 8px; color: var(--text); }
  .preview-content p { margin: 0 0 12px; }
  .preview-content a { color: var(--accent2); text-decoration: none; }
  .preview-content a:hover { text-decoration: underline; }
  .preview-content strong { color: var(--text); font-weight: 600; }
  .preview-content hr { border: none; border-top: 1px solid var(--border); margin: 20px 0; }
  .preview-content ul, .preview-content ol { padding-left: 20px; margin: 0 0 12px; }
  .preview-content li { margin-bottom: 4px; }
  .preview-content blockquote { border-left: 3px solid var(--border); padding-left: 12px; color: var(--muted); margin: 0 0 12px; }
  .preview-content code { background: var(--bg); padding: 2px 5px; border-radius: 3px; font-size: 11px; color: var(--pink); }
  .preview-content pre { background: var(--bg); padding: 12px 14px; border-radius: 6px; overflow-x: auto; margin: 12px 0; }
  .preview-content pre code { color: var(--text); padding: 0; background: none; }
  .preview-content table { border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 12px; }
  .preview-content th, .preview-content td { border: 1px solid var(--border); padding: 7px 12px; text-align: left; }
  .preview-content th { background: rgba(189,147,249,0.07); color: var(--accent); font-weight: 500; }
  .preview-content tr:nth-child(even) td { background: rgba(255,255,255,0.02); }
</style>
</head>
<body>

<header>
  <div class="logo">MD<span>Convert</span></div>
  <div class="badge">local · localhost:3200</div>  
  <div class="badge" style="margin-left:auto">v2.4 · .docx .pdf .pptx .xlsx .html .csv .eml .msg .epub .json .xml .zip .png .jpg…</div>
</header>

<main>
  <!-- LEFT -->
  <div class="left">

    <!-- Drop Zone -->
    <div class="dropzone" id="dropzone" onclick="document.getElementById('file-input').click()">
      <span class="drop-icon">⬇</span>
      <div class="drop-title">Arrastra archivos aquí</div>
      <div class="drop-sub">
        o haz clic para seleccionar<br>
        <strong>Outlook:</strong> arrastra emails directamente desde el panel de mensajes
      </div>
      <div class="file-types">
        <span class="ft ft-docx">DOCX</span>
        <span class="ft ft-pdf">PDF</span>
        <span class="ft ft-docx">PPTX</span>
        <span class="ft ft-xlsx">XLSX</span>
        <span class="ft ft-html">HTML</span>
        <span class="ft ft-eml">EML</span>
        <span class="ft ft-eml">MSG</span>
        <span class="ft ft-csv">CSV</span>
        <span class="ft ft-img">IMG</span>
        <span class="ft ft-epub">EPUB</span>
        <span class="ft ft-json">JSON</span>
        <span class="ft ft-xml">XML</span>
        <span class="ft ft-zip">ZIP</span>
      </div>
    </div>    
    <input type="file" id="file-input" multiple
      accept=".docx,.pdf,.pptx,.html,.htm,.xlsx,.csv,.eml,.msg,.epub,.json,.xml,.zip">

    <!-- Clipboard paste -->
    <button class="btn-clipboard" id="btn-clipboard" onclick="pasteFromClipboard()">
      ⎘ Pegar desde portapapeles — texto, HTML o imagen
    </button>

    <!-- Conversiones (lista unificada: en curso + historial) -->
    <div class="section-title">Conversiones</div>
    <div class="filter-wrap">
      <span class="filter-icon">⌕</span>
      <input class="filter-input" id="conv-filter" placeholder="Filtrar archivos..." oninput="renderConversions()">
    </div>
    <div class="results-list" id="conversions-list">
      <div class="empty-state">Sin archivos aún</div>
    </div>

  </div>

  <!-- RIGHT -->
  <div class="right">

    <!-- Watch Folder (Outlook) -->
    <div class="watch-card">
      <div class="watch-title">
        📁 Carpeta vigilada
      </div>
      <div class="watch-desc">
        Para <strong>Outlook</strong>: configura esta carpeta como destino de guardado automático.
        El conversor detectará nuevos archivos al instante.
      </div>
      <div class="input-row">
        <input type="text" class="input-folder" id="watch-input"
          placeholder="C:\Users\...\correos_outlook">
      </div>    
      <div class="input-row" style="align: right;">
        <button class="btn-primary" id="btn-watch-start">Iniciar</button>
      </div>
      <div class="input-row">
        <div class="watch-status" id="watch-status"></div>
      </div>  
      <div class="input-row">
        <button class="btn-open" onclick="openFolder('watch')" title="Abrir carpeta de entrada">📂 Correos</button>
      </div>
      <div class="input-row">
        <button class="btn-open" onclick="openFolder('output')" title="Abrir carpeta de salida MD">📄 Salida MD</button>
      </div>
    </div>

  </div>
</main>

<!-- Preview modal -->
<div class="preview-overlay" id="preview-overlay" onclick="closePreview(event)">
  <div class="preview-box" onclick="event.stopPropagation()">
    <div class="preview-header">
      <span class="preview-title" id="preview-title"></span>
      <button class="preview-btn" id="preview-copy-btn" onclick="copyPreviewContent()">⎘ Copiar</button>
      <a class="preview-btn" id="preview-dl-btn" href="#" download>↓ Descargar</a>
      <button class="preview-close" onclick="closePreview()">✕</button>
    </div>
    <div class="preview-content" id="preview-content"></div>
  </div>
</div>

<script>
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');

// ── Drag & Drop ──────────────────────────────────────────────────────────────

['dragenter','dragover'].forEach(e =>
  dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.add('drag-over'); })
);
['dragleave','dragend'].forEach(e =>
  dropzone.addEventListener(e, () => dropzone.classList.remove('drag-over'))
);

dropzone.addEventListener('drop', ev => {
  ev.preventDefault();
  dropzone.classList.remove('drag-over');
  const files = [...(ev.dataTransfer.files || [])];
  if (files.length) uploadFiles(files);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length) uploadFiles([...fileInput.files]);
  fileInput.value = '';
});

// ── Clipboard paste ──────────────────────────────────────────────────────────

async function pasteFromClipboard() {
  const btn = document.getElementById('btn-clipboard');
  btn.disabled = true;
  btn.classList.add('pasting');
  const label = btn.textContent;
  btn.textContent = '⌛ Leyendo portapapeles...';

  try {
    // Modern API: supports images + rich text
    if (navigator.clipboard && navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        // Image (screenshot, copy from image viewer, etc.)
        const imgType = item.types.find(t => t.startsWith('image/'));
        if (imgType) {
          const blob = await item.getType(imgType);
          const ext  = imgType === 'image/jpeg' ? '.jpg' : '.png';
          const file = new File([blob], 'clipboard' + ext, { type: imgType });
          await uploadFiles([file]);
          return;
        }
        // Rich text (HTML) — from browser, Word, email clients, etc.
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          await _sendClipboardText('html', await blob.text());
          return;
        }
        // Plain text fallback
        if (item.types.includes('text/plain')) {
          const blob = await item.getType('text/plain');
          await _sendClipboardText('text', await blob.text());
          return;
        }
      }
    }

    // Legacy fallback: readText only
    const text = await navigator.clipboard.readText();
    if (text.trim()) await _sendClipboardText('text', text);

  } catch (e) {
    const msg = e.name === 'NotAllowedError'
      ? 'Permiso denegado — permite el acceso al portapapeles en el navegador'
      : 'Error al leer portapapeles: ' + e.message;
    errorItems.unshift({ name: 'portapapeles', error: msg });
    renderConversions();
  } finally {
    btn.textContent = label;
    btn.classList.remove('pasting');
    btn.disabled = false;
  }
}

async function _sendClipboardText(type, content) {
  if (!content.trim()) {
    errorItems.unshift({ name: 'portapapeles', error: 'El portapapeles está vacío' });
    renderConversions();
    return;
  }
  const resp = await fetch('/convert-clipboard', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ type, content })
  });
  const data = await resp.json();
  if (data.error) {
    errorItems.unshift({ name: 'portapapeles', error: data.error });
  } else {
    for (const r of data) {
      if (!r.ok) errorItems.unshift({ name: r.name, error: r.error || 'Error' });
    }
  }
  await refreshConversions();
}

// ── Upload & Convert ─────────────────────────────────────────────────────────

let allFiles   = [];
let pendingMap = new Map();   // id → filename original
let errorItems = [];          // [{name, error}]

async function uploadFiles(files) {
  // Registrar pendientes
  const ids = files.map(f => {
    const id = 'p_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    pendingMap.set(id, f.name);
    return id;
  });
  renderConversions();

  const fd = new FormData();
  for (const f of files) fd.append('files', f);
  const resp = await fetch('/convert', { method: 'POST', body: fd });
  const data = await resp.json();

  // Quitar pendientes de esta tanda
  for (const id of ids) pendingMap.delete(id);

  // Acumular errores
  for (const r of data) {
    if (!r.ok) errorItems.unshift({ name: r.name, error: r.error || 'Error desconocido' });
  }

  await refreshConversions();
}

// ── Lista unificada ───────────────────────────────────────────────────────────

async function refreshConversions() {
  const resp = await fetch('/files');
  allFiles = await resp.json();
  renderConversions();
}


function renderConversions() {
  const listEl = document.getElementById('conversions-list');
  const q = (document.getElementById('conv-filter')?.value || '').toLowerCase();
  const filtered = q ? allFiles.filter(f => f.name.toLowerCase().includes(q)) : allFiles;

  let html = '';

  // En curso
  for (const [id, name] of pendingMap) {
    html += `
      <div class="result-item converting" id="${esc(id)}">
        <div class="spinner"></div>
        <div class="result-info">
          <div class="result-name">${esc(name)}</div>
          <div class="result-meta">Convirtiendo...</div>
        </div>
      </div>`;
  }

  // Errores
  for (const e of errorItems) {
    html += `
      <div class="result-item err">
        <div class="result-icon">✗</div>
        <div class="result-info">
          <div class="result-name">${esc(e.name)}</div>
          <div class="result-meta err-msg">${esc(e.error)}</div>
        </div>
      </div>`;
  }

  // Archivos del servidor
  for (const f of filtered) {
    html += `
      <div class="result-item ok">
        <div class="result-info" style="min-width:0;overflow:hidden">
          <div class="result-name" onclick="previewFile('${esc(f.name)}')"
               style="cursor:pointer" title="${esc(f.name)}">${esc(f.name)}</div>
          <div class="result-meta">${f.size_kb} KB</div>
        </div>
        <button class="btn-copy" onclick="copyFile('${esc(f.name)}', this)">⎘ Copiar</button>
        <a class="btn-dl" href="/download/${encodeURIComponent(f.name)}" download>↓ Descargar</a>
      </div>`;
  }

  if (!html) {
    html = `<div class="empty-state">${q ? 'Sin resultados' : 'Sin archivos aún'}</div>`;
  }

  listEl.innerHTML = html;
}

refreshConversions();
setInterval(refreshConversions, 4000);

// ── Copy to clipboard ────────────────────────────────────────────────────────

async function copyFile(name, btn) {
  try {
    const resp = await fetch('/preview/' + encodeURIComponent(name));
    const data = await resp.json();
    await navigator.clipboard.writeText(data.content);
    btn.classList.add('copied');
    btn.textContent = '✓';
    setTimeout(() => { btn.classList.remove('copied'); btn.textContent = '⎘'; }, 1800);
  } catch (e) {
    btn.textContent = '✗';
    setTimeout(() => { btn.textContent = '⎘'; }, 1500);
  }
}

// ── Preview modal ─────────────────────────────────────────────────────────────

let _previewContent = '';

async function previewFile(name) {
  const overlay  = document.getElementById('preview-overlay');
  const titleEl  = document.getElementById('preview-title');
  const contentEl = document.getElementById('preview-content');
  const dlBtn    = document.getElementById('preview-dl-btn');
  const copyBtn  = document.getElementById('preview-copy-btn');

  titleEl.textContent = name;
  contentEl.innerHTML = '<div class="empty-state">Cargando...</div>';
  copyBtn.textContent = '⎘ Copiar';
  copyBtn.classList.remove('copied');
  dlBtn.href = '/download/' + encodeURIComponent(name);
  dlBtn.download = name;
  overlay.classList.add('open');
  document.body.style.overflow = 'hidden';

  const resp = await fetch('/preview/' + encodeURIComponent(name));
  const data = await resp.json();
  _previewContent = data.content;
  contentEl.innerHTML = renderMd(_previewContent);
}

function renderMd(raw) {
  // Separar frontmatter YAML del cuerpo
  const fm = raw.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
  if (fm) {
    return `<div class="preview-fm">${esc(fm[1])}</div>` + marked.parse(fm[2]);
  }
  return marked.parse(raw);
}

async function copyPreviewContent() {
  const btn = document.getElementById('preview-copy-btn');
  try {
    await navigator.clipboard.writeText(_previewContent);
    btn.textContent = '✓ Copiado';
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = '⎘ Copiar'; btn.classList.remove('copied'); }, 1800);
  } catch (e) {
    btn.textContent = '✗ Error';
    setTimeout(() => { btn.textContent = '⎘ Copiar'; }, 1500);
  }
}

function closePreview(event) {
  if (event && event.target !== document.getElementById('preview-overlay')) return;
  document.getElementById('preview-overlay').classList.remove('open');
  document.body.style.overflow = '';
  _previewContent = '';
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') closePreview();
});

// ── Watch Folder ─────────────────────────────────────────────────────────────

let watching = false;

document.getElementById('btn-watch-start').addEventListener('click', async () => {
  const folder  = document.getElementById('watch-input').value.trim();
  const statusEl = document.getElementById('watch-status');
  const btn     = document.getElementById('btn-watch-start');

  if (watching) {
    await fetch('/watch/stop', { method: 'POST' });
    watching = false;
    statusEl.className = 'watch-status inactive';
    statusEl.textContent = '⏹ Vigilancia detenida';
    btn.textContent = 'Iniciar';
    btn.className = 'btn-primary';
    return;
  }

  if (!folder) {
    statusEl.className = 'watch-status inactive';
    statusEl.textContent = '⚠ Ingresa una carpeta válida';
    return;
  }

  const resp = await fetch('/watch/start', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ folder })
  });
  const data = await resp.json();

  if (data.ok) {
    watching = true;
    statusEl.className = 'watch-status active';
    statusEl.innerHTML = `<span class="pulse"></span>Vigilando: ${esc(data.folder)}`;
    btn.textContent = 'Detener';
    btn.className = 'btn-danger';
  } else {
    statusEl.className = 'watch-status inactive';
    statusEl.textContent = '✗ ' + (data.error || 'Error');
  }
});

// ── Utils ────────────────────────────────────────────────────────────────────

async function openFolder(type) {
  const folder = type === 'watch'
    ? document.getElementById('watch-input').value.trim()
    : OUTPUT_PATH;
  if (!folder) return;
  await fetch('/open-folder', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ folder })
  });
}

async function initWatcher() {
  const res  = await fetch('/watch/status');
  const data = await res.json();
  const statusEl = document.getElementById('watch-status');
  const btn  = document.getElementById('btn-watch-start');
  const input = document.getElementById('watch-input');
  if (data.active && data.folder) {
    watching = true;
    input.value = data.folder;
    statusEl.className = 'watch-status active';
    statusEl.innerHTML = '<span class="pulse"></span>Vigilando: ' + esc(data.folder);
    btn.textContent = 'Detener';
    btn.className = 'btn-danger';
  }
}

let OUTPUT_PATH = '';
fetch('/output-path').then(r => r.json()).then(d => { OUTPUT_PATH = d.path || ''; }).catch(() => {});

initWatcher();

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
</script>
</body>
</html>"""


@app.route("/watch/status")
def watch_status_route():
    return jsonify({"active": WATCH_ACTIVE, "folder": str(WATCH_DIR) if WATCH_DIR else ""})


@app.route("/output-path")
def output_path_route():
    return jsonify({"path": str(OUTPUT_DIR)})


@app.route("/open-folder", methods=["POST"])
def open_folder():
    import subprocess, platform
    data = request.json or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Sin carpeta"}), 400
    try:
        if platform.system() == "Windows":
            subprocess.Popen(["explorer", folder])
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return render_template_string(HTML)


# ─── Arranque ─────────────────────────────────────────────────────────────────

def _clear_folder(folder: Path) -> int:
    """Elimina todos los archivos de una carpeta. Retorna el número de archivos eliminados."""
    count = 0
    if folder.exists():
        for f in folder.iterdir():
            if f.is_file():
                f.unlink()
                count += 1
    return count


if __name__ == "__main__":
    print("\n  MD Converter UI")
    print(f"  → http://localhost:3200")
    print(f"  → Archivos convertidos en: {OUTPUT_DIR}")
    print("  → Ctrl+C para detener\n")

    # Limpiar carpetas al arrancar
    correos_dir = Path(DEFAULT_WATCH_DIR)
    n_correos   = _clear_folder(correos_dir)
    n_output    = _clear_folder(OUTPUT_DIR)
    if n_correos or n_output:
        print(f"  🗑  Limpieza inicial: {n_correos} correo(s), {n_output} md(s) eliminados")

    # Auto-iniciar escucha en carpeta por defecto
    if _start_watcher(DEFAULT_WATCH_DIR):
        print(f"  👁  Escuchando: {DEFAULT_WATCH_DIR}")
    else:
        print(f"  ⚠  No se pudo iniciar escucha en: {DEFAULT_WATCH_DIR}")
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:3200")).start()
    app.run(host="0.0.0.0", port=3200, debug=False)
