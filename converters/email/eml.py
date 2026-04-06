"""
Conversor de archivos .eml a Markdown.
"""

import re
from pathlib import Path

from ..html import _html_to_md_with_tables
from .thread import _split_thread
from .builders import _build_md, _seg_stem


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
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return " ".join(decoded)


def _extract_body(msg) -> str:
    """Extrae el cuerpo de un mensaje email prefiriendo HTML (para tablas)."""
    body_text = []
    body_html = []

    for part in msg.walk():
        content_type = part.get_content_type()
        disposition  = str(part.get("Content-Disposition") or "")
        if "attachment" in disposition:
            continue
        charset = part.get_content_charset() or "utf-8"
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        text = payload.decode(charset, errors="replace")
        if content_type == "text/plain":
            body_text.append(text)
        elif content_type == "text/html":
            body_html.append(text)

    if body_text:
        return "\n\n".join(body_text).strip()
    elif body_html:
        return _html_to_md_with_tables("\n".join(body_html))
    return ""


def _extract_attachments(msg) -> list[str]:
    """Retorna lista de strings con nombre y tamaño de cada adjunto."""
    result = []
    for part in msg.walk():
        disposition = str(part.get("Content-Disposition") or "")
        if "attachment" in disposition:
            filename = _decode_str(part.get_filename() or "sin_nombre")
            size_kb  = len(part.get_payload(decode=True) or b"") / 1024
            result.append(f"- `{filename}` ({size_kb:.1f} KB)")
    return result


# ─── Conversor principal ──────────────────────────────────────────────────────

def convert_eml(path: Path) -> list[tuple[str, str]]:
    """
    Convierte .eml a Markdown con YAML frontmatter.
    Si el correo contiene un hilo, retorna una lista de
    (nombre_archivo, contenido_md) — uno por cada mensaje.
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

    body        = _extract_body(msg)
    attachments = _extract_attachments(msg)

    try:
        fecha_stem = parsedate_to_datetime(date_raw).strftime("%Y-%m-%d")
    except Exception:
        fecha_stem = "0000-00-00"

    subject_slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', subject)
    subject_slug = re.sub(r'\s+', ' ', subject_slug).strip()[:80]
    base_stem    = f"{fecha_stem} — {subject_slug}"

    segments = _split_thread(body)

    if not segments:
        content = _build_md(subject, sender, to, cc, date_raw, body, attachments)
        return [(f"{base_stem}.md", content)]

    total   = len(segments)
    results = []
    for i, seg in enumerate(reversed(segments), start=1):
        seg_date    = seg.get("date")
        seg_sender  = seg.get("sender")  or sender
        seg_to      = seg.get("to")      or to
        seg_cc      = seg.get("cc")      or cc
        seg_subject = seg.get("subject") or subject
        seg_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f,;]', ' ', seg_subject)
        seg_slug    = re.sub(r' {2,}', ' ', seg_slug).strip()[:80]
        stem        = _seg_stem(seg_date, date_raw, seg_slug)
        filename    = f"{stem} — msg{i:02d} de {total}.md"
        content     = _build_md(
            seg_subject, seg_sender, seg_to, seg_cc,
            seg_date if seg_date else date_raw,
            seg["body"],
            attachments if i == total else [],
            index=i, total=total,
        )
        results.append((filename, content))

    print(f"  🧵 Hilo detectado: {total} mensajes")
    return results
