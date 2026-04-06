#!/usr/bin/env python3
"""
convert_to_md.py — Conversor universal a Markdown
Soporta: .docx, .pdf, .html, .htm, .xlsx, .csv, .eml, URLs web

Uso:
    python convert_to_md.py archivo.docx
    python convert_to_md.py reporte.pdf
    python convert_to_md.py datos.xlsx
    python convert_to_md.py pagina.html
    python convert_to_md.py correo.eml
    python convert_to_md.py https://ejemplo.com
    python convert_to_md.py carpeta/           # Convierte todos los archivos soportados
"""

import sys
import os
import argparse
from pathlib import Path


# ─── Instalador de dependencias ───────────────────────────────────────────────

def install_deps():
    """Instala dependencias si no están presentes."""
    import subprocess
    deps = [
        "mammoth",        # docx → md
        "pdfplumber",     # pdf → texto
        "html2text",      # html → md
        "pandas",         # xlsx/csv → md
        "openpyxl",       # motor Excel para pandas
        "tabulate",       # formato tabla markdown
        "requests",       # fetch URL
        "beautifulsoup4", # parseo HTML
        "extract-msg",    # .msg Outlook → datos estructurados
    ]
    print("📦 Instalando dependencias...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet"] + deps
    )
    print("✅ Dependencias listas.\n")


# ─── Conversores ──────────────────────────────────────────────────────────────

def convert_docx(path: Path) -> str:
    """Convierte .docx a Markdown usando mammoth."""
    import mammoth

    with open(path, "rb") as f:
        result = mammoth.convert_to_markdown(f)

    if result.messages:
        warnings = [m.message for m in result.messages]
        print(f"  ⚠️  Advertencias: {'; '.join(warnings)}")

    return result.value


def convert_pdf(path: Path) -> str:
    """Extrae texto de PDF con pdfplumber y lo estructura como Markdown."""
    import pdfplumber

    lines = []
    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages, 1):
            print(f"  📄 Procesando página {i}/{total}...", end="\r")

            # Texto plano
            text = page.extract_text()
            if text:
                lines.append(text.strip())

            # Tablas detectadas
            for table in page.extract_tables():
                if not table:
                    continue
                header = table[0]
                rows = table[1:]
                # Encabezado
                lines.append("\n| " + " | ".join(str(c or "") for c in header) + " |")
                lines.append("| " + " | ".join("---" for _ in header) + " |")
                for row in rows:
                    lines.append("| " + " | ".join(str(c or "") for c in row) + " |")
                lines.append("")

    print()  # salto de línea tras el progreso
    return "\n\n".join(lines)


def convert_html(source: str, is_url: bool = False) -> str:
    """Convierte HTML local o URL a Markdown."""
    import html2text
    import requests

    converter = html2text.HTML2Text()
    converter.ignore_links = False
    converter.ignore_images = False
    converter.body_width = 0  # sin wrap automático
    converter.protect_links = True
    converter.unicode_snob = True

    if is_url:
        print(f"  🌐 Descargando {source}...")
        resp = requests.get(source, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        html_content = resp.text
    else:
        with open(source, "r", encoding="utf-8", errors="replace") as f:
            html_content = f.read()

    return converter.handle(html_content)


def convert_xlsx(path: Path) -> str:
    """Convierte todas las hojas de un .xlsx a Markdown."""
    import pandas as pd

    xl = pd.ExcelFile(path)
    sections = []

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet_name)
        df = df.fillna("")
        md_table = df.to_markdown(index=False)
        sections.append(f"## {sheet_name}\n\n{md_table}")

    return "\n\n---\n\n".join(sections)


def convert_csv(path: Path) -> str:
    """Convierte .csv a tabla Markdown."""
    import pandas as pd

    df = pd.read_csv(path)
    df = df.fillna("")
    return df.to_markdown(index=False)


def _decode_str(value: str) -> str:
    """Decodifica encabezados con encoding (ej: =?UTF-8?b?...?=)."""
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
    """Extrae el cuerpo de un mensaje email como texto plano."""
    import html2text

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
        converter = html2text.HTML2Text()
        converter.ignore_links = False
        converter.body_width   = 0
        converter.unicode_snob = True
        return converter.handle("\n".join(body_html)).strip()
    return ""


def _extract_attachments(msg) -> list[str]:
    """Retorna lista de strings con nombre y tamaño de adjuntos."""
    result = []
    for part in msg.walk():
        disposition = str(part.get("Content-Disposition") or "")
        if "attachment" in disposition:
            filename = _decode_str(part.get_filename() or "sin_nombre")
            size_kb  = len(part.get_payload(decode=True) or b"") / 1024
            result.append(f"- `{filename}` ({size_kb:.1f} KB)")
    return result


def _parse_date_spanish(text: str):
    """
    Intenta parsear fechas en formatos que aparecen en correos Outlook/Gmail:
      - "Thu, Apr 2, 2026 at 6:24 PM"          (Gmail On...wrote)
      - "Tue, 24 Feb 2026 at 23:50"             (Gmail variante)
      - "jueves, 2 de abril de 2026 8:18"       (Outlook Enviado español)
      - "martes, 3 de marzo de 2026 12:06"
    Retorna datetime o None.
    """
    import re
    from datetime import datetime

    MONTHS_ES = {
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
        'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
        'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
    }
    MONTHS_EN = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }

    # Spanish: "jueves, 2 de abril de 2026 8:18" or "martes, 31 de marzo de 2026"
    m = re.search(
        r'\b(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})',
        text, re.IGNORECASE
    )
    if m:
        day, month_str, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        month = MONTHS_ES.get(month_str)
        if month:
            try:
                return datetime(year, month, day)
            except Exception:
                pass

    # English Gmail: "Apr 2, 2026" or "2 Apr 2026" or "Apr 2 2026"
    m = re.search(
        r'\b([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})',
        text, re.IGNORECASE
    )
    if m:
        month_str, day, year = m.group(1).lower()[:3], int(m.group(2)), int(m.group(3))
        month = MONTHS_EN.get(month_str)
        if month:
            try:
                return datetime(year, month, day)
            except Exception:
                pass

    m = re.search(
        r'\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})',
        text, re.IGNORECASE
    )
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
    Salta el bloque de encabezados Outlook (De:/Enviado:/Para:/Cc:/Asunto:)
    que aparece después del separador ________________________.
    Retorna (texto_contenido, meta_dict) donde meta_dict contiene:
      date, sender, to, cc, subject  (cualquiera puede ser None si no está presente)
    """
    import re
    text = text.strip()
    empty_meta = {"date": None, "sender": None, "to": None, "cc": None, "subject": None}
    if not text:
        return text, empty_meta

    lines = text.split('\n')
    HEADER_RE = re.compile(
        r'^\s*(?:De|From|Enviado(?:\s+el)?|Sent|Para|To\b|Cc|Asunto|Subject)\s*:',
        re.IGNORECASE
    )
    DE_RE      = re.compile(r'^\s*(?:De|From)\s*:\s*', re.IGNORECASE)
    ENVIADO_RE = re.compile(r'^\s*(?:Enviado(?:\s+el)?|Sent)\s*:\s*', re.IGNORECASE)
    PARA_RE    = re.compile(r'^\s*(?:Para|To)\s*:\s*', re.IGNORECASE)
    CC_RE      = re.compile(r'^\s*Cc\s*:\s*', re.IGNORECASE)
    ASUNTO_RE  = re.compile(r'^\s*(?:Asunto|Subject)\s*:\s*', re.IGNORECASE)

    first = next((l for l in lines if l.strip()), '')
    if not HEADER_RE.match(first.strip()):
        return text, empty_meta

    last_hdr = -1
    meta = {"date": None, "sender": None, "to": None, "cc": None, "subject": None}
    current_field = None  # track multiline values
    current_value = []

    def flush_field(field, value):
        """Store collected multiline value into meta."""
        val = " ".join(value).strip()
        # Clean <mailto:...> from addresses
        val = re.sub(r'\s*<mailto:[^>]+>', '', val).strip()
        val = re.sub(r'\s+>', '>', val)  # "email >" → "email>"
        if field == "sender":   meta["sender"]  = val or None
        elif field == "to":     meta["to"]      = val or None
        elif field == "cc":     meta["cc"]      = val or None
        elif field == "subject":meta["subject"] = val or None
        elif field == "date":
            d = _parse_date_spanish(val)
            if d: meta["date"] = d

    for i, line in enumerate(lines):
        stripped = line.strip()
        if HEADER_RE.match(stripped):
            # Flush previous field
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
            # Multiline continuation only for address fields (De/Para/Cc), not Asunto/Enviado
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
    - Elimina indentación común de tabs (texto citado anidado)
    - Elimina líneas de ruido: avisos de Outlook, firmas, disclaimers
    - Limpia URLs embebidas tipo <https://...> de firmas
    - Colapsa blancos excesivos
    """
    import re
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
    # Líneas que son solo una URL embebida de firma: " <https://...> "
    URL_LINE = re.compile(r'^\s*<https?://\S+>\s*$')

    lines = text.split('\n')
    clean = []
    for line in lines:
        stripped = line.strip()
        # Cortar en separador de firma "-- " o "--"
        if stripped in ('--', '-- '):
            break
        if any(p.search(stripped) for p in NOISE):
            continue
        if URL_LINE.match(line):
            continue
        clean.append(line)

    result = '\n'.join(clean)
    # Quitar tabs de indentación de Outlook (quoting anidado) línea por línea
    result = '\n'.join(line.lstrip('\t') for line in result.split('\n'))
    # Quitar indentación común de espacios restante
    result = textwrap.dedent(result).strip()
    # Colapsar 3+ líneas en blanco a 2
    result = re.sub(r'\n{3,}', '\n\n', result)
    # Limpiar <mailto:email> embebidos dejando solo el email
    result = re.sub(r'\s*<mailto:[^>]+>', '', result)
    return result


def _split_thread(body: str) -> list[dict]:
    """
    Divide el cuerpo de un email en mensajes individuales.
    Detecta los dos patrones reales de Outlook/Gmail:
      1. Separador de subrayados: ________________________
         (seguido de bloque De:/Enviado:/Para:/Cc:/Asunto:)
      2. Línea "On [fecha] [nombre] wrote:" / "escribió:"
         (puede tener <mailto:...> embebido, líneas largas)
    Fallback: bloques citados con "> "
    """
    import re

    # Normalizar saltos de línea Windows
    body = body.replace('\r\n', '\n').replace('\r', '\n')

    SEP_PATTERNS = [
        # Outlook: línea de subrayados (8 o más _)
        re.compile(r'\n[ \t]*_{8,}[ \t]*\n'),
        # Gmail/Outlook Web: "On [fecha] [persona] wrote:" — líneas largas permitidas
        re.compile(
            r'\n[ \t]*On[ \t\u202f][^\n]{5,500}?(?:wrote|escribió)\s*:\s*\n',
            re.IGNORECASE
        ),
        # "--- Original Message ---" / "--- Mensaje original ---"
        re.compile(r'\n[ \t]*-{3,}[ \t]*(?:Original Message|Mensaje original)[ \t]*-{3,}[ \t]*\n', re.IGNORECASE),
        # Bloque De:/From: standalone (sin ________________________________ previo).
        # Acepta tanto 'From: Nombre <email>' como 'From: Nombre' (sin email en la línea).
        # Lookahead: deja 'De:' al inicio del siguiente chunk para _skip_outlook_headers.
        re.compile(
            r'\n(?=(?:De|From)\s*:[^\n]{1,120}\n'
            r'(?:Enviado(?:\s+el)?|Sent)\s*:)',
            re.IGNORECASE
        ),
    ]

    # Recopilar todas las posiciones de separadores
    splits = []
    for pat in SEP_PATTERNS:
        for m in pat.finditer(body):
            splits.append((m.start(), m.end()))

    if not splits:
        # Fallback: líneas citadas con "> "
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

    # Ordenar y eliminar solapamientos, guardando el texto del separador
    import re as _re
    splits.sort()
    merged = [splits[0]]
    for s, e in splits[1:]:
        if s >= merged[-1][1]:
            merged.append((s, e))

    segments = []
    prev_end = 0

    for start, end in merged:
        chunk = body[prev_end:start]
        sep_text = body[start:end]

        # Extraer metadata del separador "On [fecha] [Nombre] <email> wrote:"
        sep_date   = _parse_date_spanish(sep_text)
        sep_sender = None
        m_on = _re.search(
            r'On[ \t\u202f][^\n]{5,500}?(?:wrote|escribió)\s*:\s*$',
            sep_text, _re.IGNORECASE | _re.MULTILINE
        )
        if m_on:
            # Find "Name <email>" immediately before "wrote:" at end of line
            line = _re.sub(r'\s*<mailto:[^>]+>', '', m_on.group(0))  # strip mailto
            # Match purely-alphabetic "Firstname Lastname <email>" before "wrote:"
            m_addr = _re.search(
                r'([A-Z][a-zA-Z-]{2,}(?:\s+[A-Z][a-zA-Z-]{2,})+\s*<[^>@\n]+@[^>\n]+>)\s*(?:wrote|escribió)\s*:',
                line
            )
            if m_addr:
                sep_sender = _re.sub(r'\s+>', '>', m_addr.group(1).strip())

        # Extract metadata from Outlook header block within this chunk
        chunk_body, outlook_meta = _skip_outlook_headers(chunk)
        cleaned = _clean_msg_segment(chunk_body)
        if cleaned:
            segments.append({
                "body": cleaned,
                "_sep_date":   sep_date,
                "_sep_sender": sep_sender,
                "_meta":       outlook_meta,  # {date, sender, to, cc, subject}
            })
        prev_end = end

    # Último chunk
    last_raw = body[prev_end:]
    last_body, last_meta = _skip_outlook_headers(last_raw)
    cleaned = _clean_msg_segment(last_body)
    if cleaned:
        segments.append({
            "body": cleaned,
            "_sep_date":   None,
            "_sep_sender": None,
            "_meta":       last_meta,
        })

    if len(segments) <= 1:
        return []

    # Asignar metadata final a cada segmento:
    # La metadata de seg[i] viene de:
    #   - su propio bloque Outlook (_meta) si existe  ← más completo
    #   - el separador "On...wrote:" que lo precede   ← solo date + sender
    # El separador sep[i] está almacenado en segments[i], pero corresponde
    # al SIGUIENTE segmento (i+1), porque sep[i] separa chunk[i] de chunk[i+1]
    result = []
    for i, seg in enumerate(segments):
        meta  = seg["_meta"]
        prev  = segments[i - 1] if i > 0 else None

        date    = meta.get("date")    or (prev["_sep_date"]   if prev else None)
        sender  = meta.get("sender")  or (prev["_sep_sender"] if prev else None)
        to_val  = meta.get("to")
        cc_val  = meta.get("cc")
        subject = meta.get("subject")

        result.append({
            "body":    seg["body"],
            "date":    date,
            "sender":  sender,
            "to":      to_val,
            "cc":      cc_val,
            "subject": subject,
        })

    return result


def _parse_addresses(raw: str) -> list[str]:
    """Convierte string de direcciones (separadas por coma o punto y coma) en lista limpia."""
    import re
    if not raw:
        return []
    # Normalizar tabs y separadores mixtos (; y ,)
    raw = raw.replace('\t', ' ')
    raw = re.sub(r'\s*[;,]\s*', '\n', raw)
    parts = [p.strip().replace('"', '') for p in raw.splitlines()]
    return [p for p in parts if p]


def _detect_direction(sender: str) -> str:
    """Determina si el correo es ENVIADO o RECIBIDO según el remitente."""
    sender_lower = sender.lower()
    sent_patterns = ["alonzo.vera", "alvera", "alonzo vera"]
    for pattern in sent_patterns:
        if pattern in sender_lower:
            return "enviado"
    return "recibido"


def _yaml_list(items: list[str], indent: str = "  ") -> str:
    """Formatea una lista Python como bloque YAML de lista."""
    if not items:
        return ""
    return "\n" + "\n".join(f"{indent}- {item}" for item in items)


def _attachment_names(raw_list: list[str]) -> list[str]:
    """Extrae solo el nombre del archivo de las entradas '- `nombre` (x KB)'."""
    import re
    names = []
    for entry in raw_list:
        m = re.search(r'`([^`]+)`', entry)
        if m:
            names.append(m.group(1))
    return names




def _load_aliases() -> list[dict]:
    """Carga reglas de alias desde contact_aliases.json (misma carpeta que el script)."""
    import json
    aliases_path = Path(__file__).parent / "contact_aliases.json"
    if not aliases_path.exists():
        return []
    try:
        with open(aliases_path, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("aliases", [])
    except Exception:
        return []


def _apply_alias(address: str, aliases: list[dict]) -> str:
    """
    Aplica reglas de alias a una dirección de correo.
    Si algún fragmento de 'match' aparece en la dirección (case-insensitive),
    devuelve el valor 'alias'. De lo contrario devuelve la dirección sin cambios.
    """
    addr_lower = address.lower()
    for rule in aliases:
        for fragment in rule.get("match", []):
            if fragment.lower() in addr_lower:
                return rule["alias"]
    return address

def _build_md(subject: str, sender: str, to: str, cc: str,
              date_raw_dt, body: str, attachments: list[str],
              index: int = None, total: int = None) -> str:
    """Construye el Markdown con YAML frontmatter según el formato estándar."""
    from email.utils import parsedate_to_datetime

    # Fecha en yyyy-MM-dd
    try:
        if isinstance(date_raw_dt, str):
            fecha = parsedate_to_datetime(date_raw_dt).strftime("%Y-%m-%d")
        else:
            fecha = date_raw_dt.strftime("%Y-%m-%d")
    except Exception:
        fecha = str(date_raw_dt)

    to_list  = _parse_addresses(to)
    cc_list  = _parse_addresses(cc)
    att_names = _attachment_names(attachments)
    direction = _detect_direction(sender)

    # Cargar aliases y aplicar a de/para/cc
    _aliases   = _load_aliases()
    sender_out = _apply_alias(sender, _aliases)
    to_list    = [_apply_alias(a, _aliases) for a in to_list]
    cc_list    = [_apply_alias(a, _aliases) for a in cc_list]

    # Asunto limpio para frontmatter: eliminar caracteres especiales
    import re as _re
    subject_clean = _re.sub(r'[<>:";/\\|?*\x00-\x1f\[\]]', ' ', subject)
    subject_clean = _re.sub(r'[,;]+', ' ', subject_clean)
    subject_clean = _re.sub(r' {2,}', ' ', subject_clean).strip()

    # Título con índice si es parte de un hilo
    title = subject
    if index is not None and total is not None and total > 1:
        title = f"{subject} [{index}/{total}]"

    # ── YAML frontmatter ──────────────────────────────────────────────
    fm_lines = [
        "---",
        f"fecha: {fecha}",
        f"de: {sender_out}",
        "para:",
    ]
    for addr in (to_list or [_apply_alias(to, _aliases)]):
        fm_lines.append(f"  - {addr}")

    fm_lines.append("cc:")
    if cc_list:
        for addr in cc_list:
            fm_lines.append(f"  - {addr}")

    fm_lines += [
        f"asunto: {subject_clean}",
        "tipo: correo",
        f"direccion: {direction}",
        "tags:",
        "  - correo",
        "adjuntos:",
    ]
    if att_names:
        for name in att_names:
            fm_lines.append(f"  - {name}")

    fm_lines.append("---")

    # ── Cuerpo ────────────────────────────────────────────────────────
    body_lines = [
        "",
        f"# {title}",
        "",
        "## Contenido",
        "",
        body.strip() if body else "*(Sin contenido)*",
    ]

    return "\n".join(fm_lines + body_lines)



def _seg_stem(seg_date, fallback_date_raw, subject_slug: str) -> str:
    """Genera el stem del nombre de archivo usando la fecha del segmento o fallback."""
    from email.utils import parsedate_to_datetime
    from datetime import datetime
    if seg_date and isinstance(seg_date, datetime):
        return f"{seg_date.strftime('%Y-%m-%d')} — {subject_slug}"
    try:
        if isinstance(fallback_date_raw, datetime):
            return f"{fallback_date_raw.strftime('%Y-%m-%d')} — {subject_slug}"
        return f"{parsedate_to_datetime(fallback_date_raw).strftime('%Y-%m-%d')} — {subject_slug}"
    except Exception:
        return f"0000-00-00 — {subject_slug}"

def convert_eml(path: Path) -> list[tuple[str, str]]:
    """
    Convierte .eml a Markdown con YAML frontmatter.
    Si el correo contiene un hilo, retorna una lista de
    (nombre_archivo, contenido_md) — uno por cada mensaje.
    """
    import email
    import re
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

    # Fecha y asunto para el nombre de archivo
    try:
        fecha_stem = parsedate_to_datetime(date_raw).strftime("%Y-%m-%d")
    except Exception:
        fecha_stem = "0000-00-00"

    subject_slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', subject)
    subject_slug = re.sub(r'\s+', ' ', subject_slug).strip()[:80]
    base_stem = f"{fecha_stem} — {subject_slug}"

    # Intentar dividir el hilo
    segments = _split_thread(body)

    if not segments:
        content = _build_md(subject, sender, to, cc, date_raw, body, attachments)
        return [(f"{base_stem}.md", content)]

    total = len(segments)
    results = []
    # Invertir: el más antiguo (último en la lista) recibe número 1
    for i, seg in enumerate(reversed(segments), start=1):
        seg_date    = seg.get("date")
        seg_sender  = seg.get("sender")  or sender
        seg_to      = seg.get("to")      or to
        seg_cc      = seg.get("cc")      or cc
        seg_subject = seg.get("subject") or subject
        seg_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f,;]', ' ', seg_subject)
        seg_slug    = re.sub(r' {2,}', ' ', seg_slug).strip()[:80]
        seg_stem    = _seg_stem(seg_date, date_raw, seg_slug)
        filename    = f"{seg_stem} — msg{i:02d} de {total}.md"
        seg_date_arg = seg_date if seg_date else date_raw
        # El mensaje más antiguo es el último segmento (i == total tras invertir)
        content  = _build_md(
            seg_subject, seg_sender, seg_to, seg_cc, seg_date_arg,
            seg["body"], attachments if i == total else [],
            index=i, total=total
        )
        results.append((filename, content))

    print(f"  🧵 Hilo detectado: {total} mensajes")
    return results


def convert_msg(path: Path) -> list[tuple[str, str]]:
    """
    Convierte .msg (formato nativo Outlook) a Markdown con YAML frontmatter.
    Reutiliza _build_md, _split_thread y helpers del conversor EML.
    Soporta hilos detectando texto citado en el cuerpo.
    """
    import re
    import extract_msg

    with extract_msg.openMsg(str(path)) as msg:
        subject  = (msg.subject or "(Sin asunto)").strip()
        sender   = (msg.sender or "").strip()
        to       = (msg.to or "").strip()
        cc       = (msg.cc or "").strip()
        date_obj = msg.date          # datetime object or None
        body     = (msg.body or "").strip()

        # Adjuntos
        att_names = []
        for att in (msg.attachments or []):
            name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "adjunto"
            size_kb = len(att.data or b"") / 1024
            att_names.append(f"- `{name}` ({size_kb:.1f} KB)")

    # Fecha para nombre de archivo y frontmatter
    if date_obj:
        fecha_stem = date_obj.strftime("%Y-%m-%d")
        date_raw   = date_obj        # _build_md acepta datetime directamente
    else:
        fecha_stem = "0000-00-00"
        date_raw   = ""

    # Slug del asunto para nombre de archivo
    subject_slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', subject)
    subject_slug = re.sub(r'\s+', ' ', subject_slug).strip()[:80]
    base_stem = f"{fecha_stem} — {subject_slug}"

    # Intentar dividir hilo
    segments = _split_thread(body)

    if not segments:
        content = _build_md(subject, sender, to, cc, date_raw, body, att_names)
        return [(f"{base_stem}.md", content)]

    total = len(segments)
    results = []
    # Invertir: el más antiguo (último en la lista) recibe número 1
    for i, seg in enumerate(reversed(segments), start=1):
        seg_date    = seg.get("date")
        seg_sender  = seg.get("sender")  or sender
        seg_to      = seg.get("to")      or to
        seg_cc      = seg.get("cc")      or cc
        seg_subject = seg.get("subject") or subject
        seg_slug    = re.sub(r'[<>:"/\\|?*\x00-\x1f,;]', ' ', seg_subject)
        seg_slug    = re.sub(r' {2,}', ' ', seg_slug).strip()[:80]
        seg_stem    = _seg_stem(seg_date, date_raw, seg_slug)
        filename    = f"{seg_stem} — msg{i:02d} de {total}.md"
        seg_date_arg = seg_date if seg_date else date_raw
        # El mensaje más antiguo es el último segmento (i == total tras invertir)
        content  = _build_md(
            seg_subject, seg_sender, seg_to, seg_cc, seg_date_arg,
            seg["body"], att_names if i == total else [],
            index=i, total=total
        )
        results.append((filename, content))

    print(f"  🧵 Hilo detectado: {total} mensajes")
    return results


# ─── Dispatcher ───────────────────────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {".docx", ".pdf", ".html", ".htm", ".xlsx", ".csv", ".eml", ".msg"}


def convert_file(source: str, output_dir: Path = None) -> Path | None:
    """Detecta el tipo y convierte. Devuelve la ruta del .md generado."""
    is_url = source.startswith("http://") or source.startswith("https://")

    if is_url:
        stem = source.split("//")[-1].split("/")[0].replace(".", "_")
        ext = ".html"
    else:
        path = Path(source)
        if not path.exists():
            print(f"❌ No encontrado: {source}")
            return None
        stem = path.stem
        ext = path.suffix.lower()

        if ext not in SUPPORTED_EXTENSIONS:
            print(f"⚠️  Formato no soportado: {ext} ({path.name})")
            return None

    # Destino del .md
    out_dir = output_dir or (Path(source).parent if not is_url else Path("."))
    out_path = out_dir / f"{stem}.md"

    print(f"\n🔄 Convirtiendo: {source}")
    print(f"   → {out_path}")

    try:
        if is_url or ext in (".html", ".htm"):
            content = convert_html(source, is_url=is_url)
        elif ext == ".docx":
            content = convert_docx(Path(source))
        elif ext == ".pdf":
            content = convert_pdf(Path(source))
        elif ext == ".xlsx":
            content = convert_xlsx(Path(source))
        elif ext == ".csv":
            content = convert_csv(Path(source))
        elif ext in (".eml", ".msg"):
            fn_convert = convert_eml if ext == ".eml" else convert_msg
            mail_results = fn_convert(Path(source))
            saved = []
            for fname, content in mail_results:
                fpath = out_dir / fname
                fpath.parent.mkdir(parents=True, exist_ok=True)
                fpath.write_text(content, encoding="utf-8")
                size_kb = fpath.stat().st_size / 1024
                print(f"   ✅ {fname} ({size_kb:.1f} KB)")
                saved.append(fpath)
            return saved[0] if saved else None
        else:
            print(f"❌ Sin conversor para {ext}")
            return None

        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
        size_kb = out_path.stat().st_size / 1024
        print(f"   ✅ Guardado ({size_kb:.1f} KB)")
        return out_path

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return None


def convert_folder(folder: Path, output_dir: Path = None) -> list[Path]:
    """Convierte todos los archivos soportados en una carpeta."""
    results = []
    files = [f for f in folder.rglob("*") if f.suffix.lower() in SUPPORTED_EXTENSIONS]

    if not files:
        print(f"⚠️  No se encontraron archivos soportados en {folder}")
        return results

    print(f"📂 Encontrados {len(files)} archivos en {folder}")
    for f in files:
        result = convert_file(str(f), output_dir)
        if result:
            results.append(result)

    return results


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Conversor universal a Markdown (.docx, .pdf, .html, .xlsx, .csv, URLs)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument(
        "sources",
        nargs="+",
        metavar="ARCHIVO_O_URL",
        help="Archivos, carpetas o URLs a convertir"
    )
    parser.add_argument(
        "-o", "--output",
        metavar="DIR",
        help="Carpeta destino para los .md generados (por defecto: misma carpeta del origen)"
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="Instalar dependencias antes de convertir"
    )

    args = parser.parse_args()

    if args.install:
        install_deps()

    output_dir = Path(args.output) if args.output else None
    converted = []

    for source in args.sources:
        path = Path(source)
        if not (source.startswith("http://") or source.startswith("https://")) and path.is_dir():
            results = convert_folder(path, output_dir)
            converted.extend(results)
        else:
            result = convert_file(source, output_dir)
            if result:
                converted.append(result)

    print(f"\n{'─'*40}")
    print(f"✨ Convertidos: {len(converted)} archivo(s)")
    for f in converted:
        print(f"   • {f}")


if __name__ == "__main__":
    main()
