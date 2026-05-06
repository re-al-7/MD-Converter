"""
Conversor de archivos .eml a Markdown.
"""

import re
from pathlib import Path

from ..html import _html_to_md_with_tables
from .thread import _split_thread
from .builders import _build_md, _seg_stem, _decode_bytes


# ─── Extracción de cuerpo y adjuntos ──────────────────────────────────────────

def _decode_str(value: str) -> str:
    """Decodifica encabezados con encoding MIME (ej: =?UTF-8?b?...?=)."""
    from email.header import decode_header
    if not value:
        return ""
    parts = decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(_decode_bytes(part, charset))
        else:
            decoded.append(part)
    return " ".join(decoded)


def _extract_plain_body(msg) -> str:
    """Extrae el cuerpo en texto plano del mensaje (primera parte text/plain)."""
    for part in msg.walk():
        if part.get_content_type() != "text/plain":
            continue
        if "attachment" in str(part.get("Content-Disposition") or ""):
            continue
        payload = part.get_payload(decode=True)
        if payload:
            return _decode_bytes(payload, part.get_content_charset())
    return ""


def _extract_html_body(msg) -> str:
    """Extrae el cuerpo HTML del mensaje (primera parte text/html)."""
    for part in msg.walk():
        if part.get_content_type() != "text/html":
            continue
        if "attachment" in str(part.get("Content-Disposition") or ""):
            continue
        payload = part.get_payload(decode=True)
        if payload:
            return _decode_bytes(payload, part.get_content_charset())
    return ""


def _extract_attachments(msg) -> list[str]:
    """Retorna lista de strings con nombre y tamaño de cada adjunto."""
    result = []
    for part in msg.walk():
        if "attachment" not in str(part.get("Content-Disposition") or ""):
            continue
        filename = _decode_str(part.get_filename() or "sin_nombre")
        size_kb  = len(part.get_payload(decode=True) or b"") / 1024
        result.append(f"- `{filename}` ({size_kb:.1f} KB)")
    return result


# ─── Conversor principal ──────────────────────────────────────────────────────

def convert_eml(path: Path) -> list[tuple[str, str]]:
    """
    Convierte .eml a Markdown con YAML frontmatter.
    Estrategia dual (igual que .msg):
      - Texto plano para detectar separadores de hilo (fiable).
      - HTML para renderizar el contenido con tablas Markdown.
    Si el split del HTML produce el mismo N de segmentos, se emparejan.
    """
    import email
    from email.utils import parsedate_to_datetime

    with open(path, "rb") as f:
        msg = email.message_from_bytes(f.read())

    subject  = _decode_str(msg.get("Subject", "(Sin asunto)"))
    sender   = _decode_str(msg.get("From", ""))
    to       = _decode_str(msg.get("To", ""))
    cc       = _decode_str(msg.get("Cc", ""))
    date_raw = msg.get("Date", "")

    plain_body = _extract_plain_body(msg)
    html_raw   = _extract_html_body(msg)
    html_md    = _html_to_md_with_tables(html_raw) if html_raw else None
    attachments = _extract_attachments(msg)

    try:
        fecha_stem = parsedate_to_datetime(date_raw).strftime("%Y-%m-%d-%H%M")
    except Exception:
        fecha_stem = "0000-00-00-0000"

    subject_slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', subject)
    subject_slug = re.sub(r'\s+', ' ', subject_slug).strip()[:80]
    base_stem    = f"{fecha_stem} — {subject_slug}"

    # Dividir hilo desde texto plano (separadores Outlook viven ahí)
    segments = _split_thread(plain_body)

    if not segments:
        # Mensaje único: HTML con tablas si disponible, si no texto plano
        body    = html_md if html_md else plain_body
        content = _build_md(subject, sender, to, cc, date_raw, body, attachments)
        return [(f"{base_stem}.md", content)]

    # Hilo: intentar dividir el HTML en paralelo para obtener cuerpos con tablas
    html_bodies = None
    if html_md:
        html_segments = _split_thread(html_md)
        if len(html_segments) == len(segments):
            html_bodies = [seg["body"] for seg in html_segments]

    total         = len(segments)
    reversed_segs = list(reversed(segments))
    reversed_html = list(reversed(html_bodies)) if html_bodies else [None] * total
    results       = []

    for i, (seg, html_body_seg) in enumerate(zip(reversed_segs, reversed_html), start=1):
        seg_date    = seg.get("date")
        seg_sender  = seg.get("sender")  or sender
        seg_to      = seg.get("to")      or to
        seg_cc      = seg.get("cc")      or cc
        seg_subject = seg.get("subject") or subject
        seg_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f,;]', ' ', seg_subject)
        seg_slug    = re.sub(r' {2,}', ' ', seg_slug).strip()[:80]
        stem        = _seg_stem(seg_date, date_raw, seg_slug)
        filename    = f"{stem} — msg{i:02d} de {total}.md"
        seg_body    = html_body_seg if html_body_seg else seg["body"]
        content     = _build_md(
            seg_subject, seg_sender, seg_to, seg_cc,
            seg_date if seg_date else date_raw,
            seg_body,
            attachments if i == total else [],
            index=i, total=total,
        )
        results.append((filename, content))

    print(f"  🧵 Hilo detectado: {total} mensajes")
    return results
