"""
Construcción de Markdown y YAML frontmatter para correos.

Exporta: _build_md, _seg_stem  (y helpers internos usados por ambos)
"""

import re
import json
from pathlib import Path

# Raíz del proyecto (converters/email/ → converters/ → raíz)
_PROJECT_ROOT = Path(__file__).parent.parent.parent


# ─── Encoding ─────────────────────────────────────────────────────────────────

def _decode_bytes(data: bytes, hint_charset: str = None) -> str:
    """
    Decodifica bytes probando el charset sugerido y luego fallbacks comunes.
    Evita caracteres de reemplazo (\ufffd) que aparecen cuando el charset
    declarado no coincide con el encoding real del contenido.
    """
    candidates = []
    if hint_charset:
        candidates.append(hint_charset)
    candidates += ['utf-8', 'windows-1252', 'latin-1']

    for enc in candidates:
        try:
            text = data.decode(enc)
            # Rechazar si hay demasiados caracteres de reemplazo
            if text.count('\ufffd') / max(len(text), 1) < 0.01:
                return text
        except (UnicodeDecodeError, LookupError):
            continue

    return data.decode('utf-8', errors='replace')


# ─── Helpers de formato ───────────────────────────────────────────────────────

def _parse_addresses(raw: str) -> list[str]:
    """Convierte un string de direcciones (separadas por , o ;) en lista limpia."""
    if not raw:
        return []
    raw = raw.replace('\t', ' ')
    raw = re.sub(r'\s*[;,]\s*', '\n', raw)
    parts = [p.strip().replace('"', '') for p in raw.splitlines()]
    return [p for p in parts if p]


def _detect_direction(sender: str) -> str:
    """Retorna 'enviado' si el remitente es Alonzo Vera, 'recibido' en caso contrario."""
    sender_lower = sender.lower()
    for pattern in ("alonzo.vera", "alvera", "alonzo vera"):
        if pattern in sender_lower:
            return "enviado"
    return "recibido"


def _attachment_names(raw_list: list[str]) -> list[str]:
    """Extrae el nombre de archivo de entradas con formato '- `nombre` (x KB)'."""
    names = []
    for entry in raw_list:
        m = re.search(r'`([^`]+)`', entry)
        if m:
            names.append(m.group(1))
    return names


# ─── Alias de contactos ───────────────────────────────────────────────────────

def _load_aliases() -> list[dict]:
    """Carga reglas de alias desde contact_aliases.json en la raíz del proyecto."""
    aliases_path = _PROJECT_ROOT / "contact_aliases.json"
    if not aliases_path.exists():
        return []
    try:
        with open(aliases_path, encoding="utf-8") as f:
            return json.load(f).get("aliases", [])
    except Exception:
        return []


def _apply_alias(address: str, aliases: list[dict]) -> str:
    """Aplica reglas de alias a una dirección; devuelve el alias si hay match."""
    addr_lower = address.lower()
    for rule in aliases:
        for fragment in rule.get("match", []):
            if fragment.lower() in addr_lower:
                return rule["alias"]
    return address


# ─── Construcción de Markdown ─────────────────────────────────────────────────

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

    aliases   = _load_aliases()
    to_list   = [_apply_alias(a, aliases) for a in _parse_addresses(to)]
    cc_list   = [_apply_alias(a, aliases) for a in _parse_addresses(cc)]
    att_names = _attachment_names(attachments)
    sender_out = _apply_alias(sender, aliases)
    direction  = _detect_direction(sender)

    subject_clean = re.sub(r'[<>:";/\\|?*\x00-\x1f\[\]]', ' ', subject)
    subject_clean = re.sub(r'[,;]+', ' ', subject_clean)
    subject_clean = re.sub(r' {2,}', ' ', subject_clean).strip()

    title = subject
    if index is not None and total is not None and total > 1:
        title = f"{subject} [{index}/{total}]"

    # ── YAML frontmatter ──────────────────────────────────────────────
    fm = [
        "---",
        f"fecha: {fecha}",
        f"de: {sender_out}",
        "para:",
    ]
    for addr in (to_list or [_apply_alias(to, aliases)]):
        fm.append(f"  - {addr}")

    fm.append("cc:")
    for addr in cc_list:
        fm.append(f"  - {addr}")

    fm += [
        f"asunto: {subject_clean}",
        "tipo: correo",
        f"direccion: {direction}",
        "tags:",
        "  - correo",
        "adjuntos:",
    ]
    for name in att_names:
        fm.append(f"  - {name}")

    fm.append("---")

    # ── Cuerpo ────────────────────────────────────────────────────────
    body_md = [
        "",
        f"# {title}",
        "",
        "## Contenido",
        "",
        body.strip() if body else "*(Sin contenido)*",
    ]

    return "\n".join(fm + body_md)


def _seg_stem(seg_date, fallback_date_raw, subject_slug: str) -> str:
    """Genera el stem del nombre de archivo usando la fecha del segmento o fallback."""
    from email.utils import parsedate_to_datetime
    from datetime import datetime

    if seg_date and isinstance(seg_date, datetime):
        return f"{seg_date.strftime('%Y-%m-%d-%H%M')} — {subject_slug}"
    try:
        if isinstance(fallback_date_raw, datetime):
            return f"{fallback_date_raw.strftime('%Y-%m-%d-%H%M')} — {subject_slug}"
        return f"{parsedate_to_datetime(fallback_date_raw).strftime('%Y-%m-%d-%H%M')} — {subject_slug}"
    except Exception:
        return f"0000-00-00-0000 — {subject_slug}"
