"""
Conversor de archivos .msg (Outlook) a Markdown.
"""

import re
from pathlib import Path

from ..html import _html_to_md_with_tables
from .thread import _split_thread
from .builders import _build_md, _seg_stem, _decode_bytes


def convert_msg(path: Path) -> list[tuple[str, str]]:
    """
    Convierte .msg (formato nativo Outlook) a Markdown con YAML frontmatter.
    Estrategia dual:
      - Divide hilos usando el cuerpo de texto plano (separadores Outlook fiables).
      - Renderiza el contenido desde el HTML body (preserva tablas Markdown).
      Si ambos splits coinciden en N segmentos, se emparejan metadata+cuerpo.
      Si no coinciden, se usa el texto plano como fallback.
    """
    import extract_msg

    with extract_msg.openMsg(str(path)) as msg:
        subject       = (msg.subject or "(Sin asunto)").strip()
        sender        = (msg.sender  or "").strip()
        to            = (msg.to      or "").strip()
        cc            = (msg.cc      or "").strip()
        date_obj      = msg.date
        plain_body    = (msg.body    or "").strip()
        html_body_raw = getattr(msg, 'htmlBody', None)

        att_names = []
        for att in (msg.attachments or []):
            name    = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "adjunto"
            size_kb = len(att.data or b"") / 1024
            att_names.append(f"- `{name}` ({size_kb:.1f} KB)")

    # Fecha y stem base
    if date_obj:
        fecha_stem = date_obj.strftime("%Y-%m-%d")
        date_raw   = date_obj
    else:
        fecha_stem = "0000-00-00"
        date_raw   = ""

    subject_slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', subject)
    subject_slug = re.sub(r'\s+', ' ', subject_slug).strip()[:80]
    base_stem    = f"{fecha_stem} — {subject_slug}"

    # Decodificar y convertir HTML body una sola vez (con fallback de encoding)
    if html_body_raw and isinstance(html_body_raw, bytes):
        html_body_raw = _decode_bytes(html_body_raw)
    html_md = _html_to_md_with_tables(html_body_raw) if html_body_raw else None

    # Dividir hilo desde texto plano (separadores Outlook viven ahí)
    segments = _split_thread(plain_body)

    if not segments:
        # Mensaje único: usar HTML body para conservar tablas
        body    = html_md if html_md else plain_body
        content = _build_md(subject, sender, to, cc, date_raw, body, att_names)
        return [(f"{base_stem}.md", content)]

    # Hilo: intentar dividir también el HTML para obtener cuerpos con tablas.
    # La metadata siempre viene del split de texto plano (más fiable).
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
            att_names if i == total else [],
            index=i, total=total,
        )
        results.append((filename, content))

    print(f"  🧵 Hilo detectado: {total} mensajes")
    return results
