#!/usr/bin/env python3
"""
converter_ui.py — UI web local para el conversor a Markdown
Ejecutar: python converter_ui.py
Luego abrir: http://localhost:5000
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

from convert_to_md import convert_docx, convert_pdf, convert_html, convert_xlsx, convert_csv, convert_eml, convert_msg, convert_image, IMAGE_EXTENSIONS

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

SUPPORTED = {".docx", ".pdf", ".html", ".htm", ".xlsx", ".csv", ".eml", ".msg"} | IMAGE_EXTENSIONS

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
            elif ext == ".pdf":            content = convert_pdf(src)
            elif ext in (".html", ".htm"): content = convert_html(str(src))
            elif ext == ".xlsx":           content = convert_xlsx(src)
            elif ext == ".csv":            content = convert_csv(src)
            elif ext in IMAGE_EXTENSIONS:  content = convert_image(src)
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


# ─── UI ───────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MD Converter</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@400;700;800&display=swap" rel="stylesheet">
<style>
  :root {
    --bg:      #0b0d11;
    --surface: #13161d;
    --border:  #1f2330;
    --accent:  #00e5a0;
    --accent2: #0066ff;
    --warn:    #ff4d6d;
    --text:    #e2e8f0;
    --muted:   #4a5568;
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
    grid-template-columns: 1fr 1fr 280px;
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
    background: radial-gradient(ellipse at 50% 0%, rgba(0,229,160,0.04) 0%, transparent 70%);
    pointer-events: none;
  }

  .dropzone.drag-over {
    border-color: var(--accent);
    background: rgba(0,229,160,0.05);
    transform: scale(1.005);
  }

  .dropzone.drag-over::before {
    background: radial-gradient(ellipse at 50% 0%, rgba(0,229,160,0.12) 0%, transparent 70%);
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
    background: rgba(0,229,160,0.1);
    border: 1px solid rgba(0,229,160,0.25);
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

  .btn-dl:hover {
    background: rgba(0,229,160,0.2);
    border-color: var(--accent);
  }

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

  .watch-desc strong { color: #c084fc; }

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
    color: #000;
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
    background: rgba(0,229,160,0.1);
    border: 1px solid rgba(0,229,160,0.2);
    color: var(--accent);
  }

  .watch-status.inactive {
    display: block;
    background: rgba(255,77,109,0.1);
    border: 1px solid rgba(255,77,109,0.2);
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

  /* ── CENTER PANEL (Preview) ── */
  .center {
    display: flex;
    flex-direction: column;
    height: calc(100vh - 61px);
    border-right: 1px solid var(--border);
    overflow: hidden;
  }

  .preview-empty {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 10px;
    color: var(--muted);
    font-size: 12px;
  }

  .preview-empty-icon {
    font-size: 36px;
    opacity: 0.2;
    line-height: 1;
  }

  .preview-panel {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .preview-header {
    padding: 10px 16px;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    gap: 10px;
    flex-shrink: 0;
    background: var(--surface);
  }

  .preview-label {
    font-family: var(--display);
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--muted);
    flex-shrink: 0;
  }

  .preview-filename {
    flex: 1;
    font-size: 11px;
    color: var(--accent);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .btn-copy {
    background: rgba(0,229,160,0.08);
    border: 1px solid rgba(0,229,160,0.2);
    color: var(--accent);
    padding: 4px 11px;
    border-radius: 4px;
    cursor: pointer;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    white-space: nowrap;
    transition: all 0.15s;
    flex-shrink: 0;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    gap: 5px;
  }

  .btn-copy:hover {
    background: rgba(0,229,160,0.18);
    border-color: var(--accent);
  }

  .btn-copy.copied {
    background: rgba(0,229,160,0.25);
    border-color: var(--accent);
  }

  .preview-content {
    flex: 1;
    overflow-y: auto;
    padding: 20px 24px;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: var(--mono);
    font-size: 12px;
    line-height: 1.75;
    color: var(--text);
    background: var(--bg);
    margin: 0;
    border: none;
    outline: none;
  }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
</style>
</head>
<body>

<header>
  <div class="logo">MD<span>Convert</span></div>
  <div class="badge">local · localhost:5000</div>
  <div class="badge" style="margin-left:auto">v2.2 · .docx .pdf .xlsx .html .csv .eml .msg .png .jpg…</div>
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
        <span class="ft ft-xlsx">XLSX</span>
        <span class="ft ft-html">HTML</span>
        <span class="ft ft-eml">EML</span>
        <span class="ft ft-eml">MSG</span>
        <span class="ft ft-csv">CSV</span>
        <span class="ft ft-img">IMG</span>
      </div>
    </div>
    <input type="file" id="file-input" multiple
      accept=".docx,.pdf,.html,.htm,.xlsx,.csv,.eml,.msg,.jpg,.jpeg,.png,.bmp,.tiff,.tif,.webp">

    <!-- Results -->
    <div class="section-title">Conversiones</div>
    <div class="results-list" id="results"></div>

  </div>

  <!-- CENTER - Preview -->
  <div class="center">

    <div class="preview-empty" id="preview-empty">
      <div class="preview-empty-icon">◈</div>
      <span>Convierte un archivo para ver el preview</span>
      <span style="font-size:11px;opacity:0.6">o haz clic en un archivo del historial</span>
    </div>

    <div class="preview-panel" id="preview-panel" style="display:none">
      <div class="preview-header">
        <span class="preview-label">Preview</span>
        <span class="preview-filename" id="preview-filename"></span>
        <button class="btn-copy" id="btn-copy" onclick="copyMd()">⎘ Copiar MD</button>
        <a class="btn-copy" id="btn-preview-dl" href="#" download>↓ Descargar</a>
      </div>
      <pre class="preview-content" id="preview-content"></pre>
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
        <button class="btn-primary" id="btn-watch-start">Iniciar</button>
      </div>
      <div class="watch-status" id="watch-status"></div>
      <div style="display:flex;gap:6px;margin-top:8px">
        <button class="btn-open" onclick="openFolder('watch')" title="Abrir carpeta de entrada">📂 Correos</button>
        <button class="btn-open" onclick="openFolder('output')" title="Abrir carpeta de salida MD">📄 Salida MD</button>
      </div>
    </div>

    <!-- Recent files -->
    <div class="section-title">Archivos generados</div>
    <div class="history-list" id="history">
      <div class="empty-state">Sin archivos aún</div>
    </div>

  </div>
</main>

<script>
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const results  = document.getElementById('results');
const history  = document.getElementById('history');

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

// ── Upload & Convert ─────────────────────────────────────────────────────────

async function uploadFiles(files) {
  const fd = new FormData();
  for (const f of files) fd.append('files', f);

  // Pending items
  const itemIds = files.map(f => addPending(f.name));

  const resp = await fetch('/convert', { method: 'POST', body: fd });
  const data = await resp.json();

  // Remove pending
  itemIds.forEach(id => document.getElementById(id)?.remove());

  // Show results
  for (const r of data) addResult(r);

  refreshHistory();
}

function addPending(name) {
  const id = 'p_' + Date.now() + Math.random();
  const el = document.createElement('div');
  el.className = 'result-item converting';
  el.id = id;
  el.innerHTML = `
    <div class="spinner"></div>
    <div class="result-info">
      <div class="result-name">${esc(name)}</div>
      <div class="result-meta">Convirtiendo...</div>
    </div>`;
  results.prepend(el);
  return id;
}

function addResult(r) {
  const el = document.createElement('div');
  el.className = `result-item ${r.ok ? 'ok' : 'err'}`;
  if (r.ok) {
    el.style.cursor = 'pointer';
    el.innerHTML = `
      <div class="result-icon">✓</div>
      <div class="result-info">
        <div class="result-name">${esc(r.name)}</div>
        <div class="result-meta">Convertido correctamente</div>
      </div>
      <a class="btn-dl" href="/download/${encodeURIComponent(r.name)}" download>↓ Descargar</a>`;
    el.addEventListener('click', e => {
      if (!e.target.closest('.btn-dl')) loadPreview(r.name);
    });
    loadPreview(r.name);
  } else {
    el.innerHTML = `
      <div class="result-icon">✗</div>
      <div class="result-info">
        <div class="result-name">${esc(r.name)}</div>
        <div class="result-meta err-msg">${esc(r.error || 'Error desconocido')}</div>
      </div>`;
  }
  results.prepend(el);
}

// ── History ──────────────────────────────────────────────────────────────────

async function refreshHistory() {
  const resp = await fetch('/files');
  const files = await resp.json();
  history.innerHTML = files.length === 0
    ? '<div class="empty-state">Sin archivos aún</div>'
    : files.map(f => `
      <div class="hist-item" style="cursor:pointer" onclick="loadPreview('${esc(f.name).replace(/'/g,"\\'")}')">
        <span class="hist-name" title="${esc(f.name)}">📄 ${esc(f.name)}</span>
        <span class="hist-size">${f.size_kb} KB</span>
        <a class="hist-dl" href="/download/${encodeURIComponent(f.name)}" download title="Descargar" onclick="event.stopPropagation()">⬇</a>
      </div>`).join('');
}

// ── Preview ───────────────────────────────────────────────────────────────────

async function loadPreview(filename) {
  const resp = await fetch(`/preview/${encodeURIComponent(filename)}`);
  if (!resp.ok) return;
  const text = await resp.text();

  document.getElementById('preview-filename').textContent = filename;
  document.getElementById('preview-content').textContent = text;
  document.getElementById('btn-preview-dl').href = `/download/${encodeURIComponent(filename)}`;
  document.getElementById('btn-preview-dl').download = filename;
  document.getElementById('btn-copy').textContent = '⎘ Copiar MD';
  document.getElementById('btn-copy').classList.remove('copied');

  document.getElementById('preview-empty').style.display = 'none';
  document.getElementById('preview-panel').style.display = 'flex';
}

async function copyMd() {
  const text = document.getElementById('preview-content').textContent;
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
  }
  const btn = document.getElementById('btn-copy');
  btn.textContent = '✓ Copiado';
  btn.classList.add('copied');
  setTimeout(() => { btn.textContent = '⎘ Copiar MD'; btn.classList.remove('copied'); }, 2000);
}

refreshHistory();
setInterval(refreshHistory, 4000);

// ── Watch Folder ─────────────────────────────────────────────────────────────

let watching = false;

document.getElementById('btn-watch-start').addEventListener('click', async () => {
  const folder = document.getElementById('watch-input').value.trim();
  const statusEl = document.getElementById('watch-status');
  const btn = document.getElementById('btn-watch-start');

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

// Auto-start watch on load and sync UI state
async function initWatcher() {
  const res = await fetch('/watch/status');
  const data = await res.json();
  const statusEl = document.getElementById('watch-status');
  const btn = document.getElementById('btn-watch-start');
  const input = document.getElementById('watch-input');
  if (data.active && data.folder) {
    watching = true;
    if (data.folder) input.value = data.folder;
    statusEl.className = 'watch-status active';
    statusEl.innerHTML = '<span class="pulse"></span>Vigilando: ' + esc(data.folder);
    btn.textContent = 'Detener';
    btn.className = 'btn-danger';
  }
}

let OUTPUT_PATH = '';
fetch('/files').then(r => r.json()).then(() => {});
// Get output path from first file or derive it
OUTPUT_PATH = '';
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


@app.route("/preview/<path:filename>")
def preview(filename):
    fp = OUTPUT_DIR / filename
    if not fp.exists():
        return "Archivo no encontrado", 404
    return fp.read_text(encoding="utf-8"), 200, {"Content-Type": "text/plain; charset=utf-8"}


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

if __name__ == "__main__":
    print("\n  MD Converter UI")
    print(f"  → http://localhost:5000")
    print(f"  → Archivos convertidos en: {OUTPUT_DIR}")
    print("  → Ctrl+C para detener\n")
    # Auto-iniciar escucha en carpeta por defecto
    if _start_watcher(DEFAULT_WATCH_DIR):
        print(f"  👁  Escuchando: {DEFAULT_WATCH_DIR}")
    else:
        print(f"  ⚠  No se pudo iniciar escucha en: {DEFAULT_WATCH_DIR}")
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5000")).start()
    app.run(host="0.0.0.0", port=5000, debug=False)
