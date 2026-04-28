"""
Detección y separación de hilos de correo.

Exporta: _parse_date_spanish, _skip_outlook_headers, _clean_msg_segment, _split_thread
"""

import re


def _parse_date_spanish(text: str):
    """
    Parsea fechas en formatos que aparecen en correos Outlook/Gmail:
      - "Thu, Apr 2, 2026 at 6:24 PM"        (Gmail On...wrote)
      - "jueves, 2 de abril de 2026 8:18"     (Outlook Enviado español)
    Retorna datetime o None.
    """
    from datetime import datetime

    MONTHS_ES = {
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
        'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
        'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12,
    }
    MONTHS_EN = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    }

    # Español: "2 de abril de 2026"
    m = re.search(r'\b(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})', text, re.IGNORECASE)
    if m:
        day, month_str, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        month = MONTHS_ES.get(month_str)
        if month:
            try:
                return datetime(year, month, day)
            except Exception:
                pass

    # Inglés: "Apr 2, 2026"
    m = re.search(r'\b([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})', text, re.IGNORECASE)
    if m:
        month_str, day, year = m.group(1).lower()[:3], int(m.group(2)), int(m.group(3))
        month = MONTHS_EN.get(month_str)
        if month:
            try:
                return datetime(year, month, day)
            except Exception:
                pass

    # Inglés: "2 Apr 2026"
    m = re.search(r'\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})', text, re.IGNORECASE)
    if m:
        day, month_str, year = int(m.group(1)), m.group(2).lower()[:3], int(m.group(3))
        month = MONTHS_EN.get(month_str)
        if month:
            try:
                return datetime(year, month, day)
            except Exception:
                pass

    return None


def _skip_outlook_headers(text: str) -> tuple:
    """
    Salta el bloque de encabezados Outlook al inicio de un segmento.
    Acepta tanto formato texto plano (De:) como html2text con negritas (**De:**).
    Retorna (cuerpo, meta_dict) donde meta_dict contiene date, sender, to, cc, subject.
    """
    text = text.strip()
    empty_meta = {"date": None, "sender": None, "to": None, "cc": None, "subject": None}
    if not text:
        return text, empty_meta

    lines = text.split('\n')

    # Acepta "De:" (texto plano) y "**De:**" (html2text con negritas)
    _F = r'(?:\*\*)?(?:De|From|Enviado(?:\s+el)?|Sent|Para|To\b|Cc|Asunto|Subject)(?::\*\*|\s*:)'
    HEADER_RE  = re.compile(r'^\s*' + _F, re.IGNORECASE)
    DE_RE      = re.compile(r'^\s*(?:\*\*)?(?:De|From)(?::\*\*|\s*:)\s*', re.IGNORECASE)
    ENVIADO_RE = re.compile(r'^\s*(?:\*\*)?(?:Enviado(?:\s+el)?|Sent)(?::\*\*|\s*:)\s*', re.IGNORECASE)
    PARA_RE    = re.compile(r'^\s*(?:\*\*)?(?:Para|To)(?::\*\*|\s*:)\s*', re.IGNORECASE)
    CC_RE      = re.compile(r'^\s*(?:\*\*)?Cc(?::\*\*|\s*:)\s*', re.IGNORECASE)
    ASUNTO_RE  = re.compile(r'^\s*(?:\*\*)?(?:Asunto|Subject)(?::\*\*|\s*:)\s*', re.IGNORECASE)

    first = next((l for l in lines if l.strip()), '')
    if not HEADER_RE.match(first.strip()):
        return text, empty_meta

    last_hdr = -1
    meta = {"date": None, "sender": None, "to": None, "cc": None, "subject": None}
    current_field = None
    current_value = []

    def flush_field(field, value):
        val = " ".join(value).strip()
        val = re.sub(r'\s*<mailto:[^>]+>', '', val).strip()
        val = re.sub(r'\s+>', '>', val)
        if field == "sender":    meta["sender"]  = val or None
        elif field == "to":      meta["to"]      = val or None
        elif field == "cc":      meta["cc"]      = val or None
        elif field == "subject": meta["subject"] = val or None
        elif field == "date":
            d = _parse_date_spanish(val)
            if d:
                meta["date"] = d

    for i, line in enumerate(lines):
        stripped = line.strip()
        if HEADER_RE.match(stripped):
            if current_field:
                flush_field(current_field, current_value)
            last_hdr = i
            current_value = []
            if DE_RE.match(stripped):
                current_field = "sender"
                current_value = [DE_RE.sub("", stripped)]
            elif ENVIADO_RE.match(stripped):
                current_field = "date"
                current_value = [ENVIADO_RE.sub("", stripped)]
            elif PARA_RE.match(stripped):
                current_field = "to"
                current_value = [PARA_RE.sub("", stripped)]
            elif CC_RE.match(stripped):
                current_field = "cc"
                current_value = [CC_RE.sub("", stripped)]
            elif ASUNTO_RE.match(stripped):
                current_field = "subject"
                current_value = [ASUNTO_RE.sub("", stripped)]
            else:
                current_field = None
        elif last_hdr >= 0 and current_field and stripped:
            # Continuación multilínea solo para campos de dirección
            if current_field in ("to", "cc") and line.startswith((' ', '\t')):
                current_value.append(stripped)
            else:
                flush_field(current_field, current_value)
                current_field = None
                break
        elif last_hdr >= 0 and stripped:
            flush_field(current_field, current_value)
            current_field = None
            break

    if current_field:
        flush_field(current_field, current_value)

    body = '\n'.join(lines[last_hdr + 1:]).strip() if last_hdr >= 0 else text
    return body, meta


def _clean_msg_segment(text: str) -> str:
    """
    Limpia un segmento de mensaje:
    - Elimina líneas de ruido (avisos Outlook, disclaimers, firmas)
    - Quita indentación de tabs (texto citado anidado)
    - Colapsa líneas en blanco excesivas
    """
    import textwrap

    text = text.strip()
    if not text:
        return ''

    NOISE = [
        re.compile(r'No suele recibir correo electrónico de', re.IGNORECASE),
        re.compile(r'Por qué es esto importante', re.IGNORECASE),
        re.compile(r'https?://aka\.ms/', re.IGNORECASE),
        re.compile(r'This e-?mail and any attachments? are confidential', re.IGNORECASE),
        re.compile(r'Imprime sólo si es necesario', re.IGNORECASE),
        re.compile(r'Esta comunicación puede contener información', re.IGNORECASE),
        re.compile(r'This transmission may contain', re.IGNORECASE),
        re.compile(r'do not constitute or imply consent', re.IGNORECASE),
        re.compile(r'Kindly note that any communications', re.IGNORECASE),
    ]
    URL_LINE = re.compile(r'^\s*<https?://\S+>\s*$')

    lines = text.split('\n')
    clean = []
    for line in lines:
        stripped = line.strip()
        if stripped in ('--', '-- '):
            break
        if any(p.search(stripped) for p in NOISE):
            continue
        if URL_LINE.match(line):
            continue
        clean.append(line)

    result = '\n'.join(clean)
    result = '\n'.join(line.lstrip('\t') for line in result.split('\n'))
    result = textwrap.dedent(result).strip()
    result = re.sub(r'\n{3,}', '\n\n', result)
    result = re.sub(r'\s*<mailto:[^>]+>', '', result)
    # Eliminar referencias a imágenes embebidas cid: (inútiles fuera del cliente de correo)
    result = re.sub(r'!\[[^\]]*\]\(cid:[^)]+\)', '', result)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip()


def _split_thread(body: str) -> list[dict]:
    """
    Divide el cuerpo de un email en mensajes individuales detectando separadores:
      1. Subrayados Outlook: ________________________
      2. Gmail/Outlook Web: "On [fecha] [persona] wrote:"
      3. "--- Original Message ---"
      4. Bloque De:/From: + Enviado:/Sent: (texto plano)
      5. Bloque **De:**/**From:** + **Enviado:**/**Sent:** (html2text con negritas)
    """
    body = body.replace('\r\n', '\n').replace('\r', '\n')

    SEP_PATTERNS = [
        # Outlook: línea de subrayados (8 o más _)
        re.compile(r'\n[ \t]*_{8,}[ \t]*\n'),
        # Gmail/Outlook Web: "On [fecha] [persona] wrote:"
        re.compile(
            r'\n[ \t]*On[ \t\u202f][^\n]{5,500}?(?:wrote|escribió)\s*:\s*\n',
            re.IGNORECASE
        ),
        # "--- Original Message ---" / "--- Mensaje original ---"
        re.compile(
            r'\n[ \t]*-{3,}[ \t]*(?:Original Message|Mensaje original)[ \t]*-{3,}[ \t]*\n',
            re.IGNORECASE
        ),
        # Bloque De:/From: standalone (texto plano)
        re.compile(
            r'\n(?=(?:De|From)\s*:[^\n]{1,120}\n(?:Enviado(?:\s+el)?|Sent)\s*:)',
            re.IGNORECASE
        ),
        # Variante html2text: **From:** + **Sent:** (negritas Markdown)
        re.compile(
            r'\n(?=[ \t]*\*\*(?:De|From):\*\*[^\n]{0,200}\n[ \t]*\*\*(?:Enviado(?:\s+el)?|Sent):\*\*)',
            re.IGNORECASE
        ),
    ]

    splits = []
    for pat in SEP_PATTERNS:
        for m in pat.finditer(body):
            splits.append((m.start(), m.end()))

    if not splits:
        # Fallback: bloques citados con "> "
        lines = body.split('\n')
        segments, current, in_quote = [], [], False
        for line in lines:
            is_q = line.startswith('>')
            if is_q != in_quote:
                chunk = '\n'.join(current).strip()
                if chunk:
                    segments.append({"body": _clean_msg_segment(chunk)})
                current, in_quote = [], is_q
            current.append(line.lstrip('> ') if is_q else line)
        chunk = '\n'.join(current).strip()
        if chunk:
            segments.append({"body": _clean_msg_segment(chunk)})
        return [s for s in segments if s["body"]] if len(segments) > 1 else []

    splits.sort()
    merged = [splits[0]]
    for s, e in splits[1:]:
        if s >= merged[-1][1]:
            merged.append((s, e))

    segments = []
    prev_end  = 0

    for start, end in merged:
        chunk    = body[prev_end:start]
        sep_text = body[start:end]

        sep_date   = _parse_date_spanish(sep_text)
        sep_sender = None
        m_on = re.search(
            r'On[ \t\u202f][^\n]{5,500}?(?:wrote|escribió)\s*:\s*$',
            sep_text, re.IGNORECASE | re.MULTILINE
        )
        if m_on:
            line = re.sub(r'\s*<mailto:[^>]+>', '', m_on.group(0))
            m_addr = re.search(
                r'([A-Z][a-zA-Z-]{2,}(?:\s+[A-Z][a-zA-Z-]{2,})+\s*<[^>@\n]+@[^>\n]+>)\s*(?:wrote|escribió)\s*:',
                line
            )
            if m_addr:
                sep_sender = re.sub(r'\s+>', '>', m_addr.group(1).strip())

        chunk_body, outlook_meta = _skip_outlook_headers(chunk)
        cleaned = _clean_msg_segment(chunk_body)
        if cleaned:
            segments.append({
                "body":        cleaned,
                "_sep_date":   sep_date,
                "_sep_sender": sep_sender,
                "_meta":       outlook_meta,
            })
        prev_end = end

    # Último chunk
    last_raw  = body[prev_end:]
    last_body, last_meta = _skip_outlook_headers(last_raw)
    cleaned   = _clean_msg_segment(last_body)
    if cleaned:
        segments.append({
            "body":        cleaned,
            "_sep_date":   None,
            "_sep_sender": None,
            "_meta":       last_meta,
        })

    if len(segments) <= 1:
        return []

    # Asignar metadata final: el bloque Outlook propio tiene prioridad;
    # si no, se usa la del separador "On...wrote:" que PRECEDE al segmento.
    result = []
    for i, seg in enumerate(segments):
        meta   = seg["_meta"]
        prev   = segments[i - 1] if i > 0 else None
        result.append({
            "body":    seg["body"],
            "date":    meta.get("date")    or (prev["_sep_date"]   if prev else None),
            "sender":  meta.get("sender")  or (prev["_sep_sender"] if prev else None),
            "to":      meta.get("to"),
            "cc":      meta.get("cc"),
            "subject": meta.get("subject"),
        })

    return result
