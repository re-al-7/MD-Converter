"""
Construcción de Markdown y YAML frontmatter para correos.

Exporta: _build_md, _seg_stem, _find_template  (y helpers internos usados por ambos)
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
    Evita caracteres de reemplazo (�) que aparecen cuando el charset
    declarado no coincide con el encoding real del contenido.
    """
    candidates = []
    if hint_charset:
        candidates.append(hint_charset)
    candidates += ['utf-8', 'windows-1252', 'latin-1']

    for enc in candidates:
        try:
            text = data.decode(enc)
            if text.count('�') / max(len(text), 1) < 0.01:
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


# ─── Sistema de templates ─────────────────────────────────────────────────────

_DEFAULT_TEMPLATE = {
    "name": "default",
    "match": {},
    "frontmatter": [
        "fecha",
        "de",
        "para",
        "cc",
        "asunto",
        {"name": "tipo", "value": "correo"},
        "direccion",
        {"name": "tags", "value": ["correo"]},
        "adjuntos",
    ],
    "body_header": "## Contenido",
    "filename_format": "{date} — {subject}",
}


def _load_templates() -> list[dict]:
    """Carga templates desde email_templates.json en la raíz del proyecto."""
    templates_path = _PROJECT_ROOT / "email_templates.json"
    if not templates_path.exists():
        return []
    try:
        with open(templates_path, encoding="utf-8") as f:
            return json.load(f).get("templates", [])
    except Exception:
        return []


def _find_template(sender: str, to: str, cc: str, subject: str) -> dict:
    """Retorna el primer template que hace match con los metadatos del correo, o el default."""
    templates = _load_templates()
    sender_l  = (sender  or "").lower()
    to_l      = (to      or "").lower()
    cc_l      = (cc      or "").lower()
    subject_l = (subject or "").lower()

    for tmpl in templates:
        match = tmpl.get("match", {})
        if match.get("from_contains")    and match["from_contains"].lower()    not in sender_l:  continue
        if match.get("to_contains")      and match["to_contains"].lower()      not in to_l:      continue
        if match.get("cc_contains")      and match["cc_contains"].lower()      not in cc_l:      continue
        if match.get("subject_contains") and match["subject_contains"].lower() not in subject_l: continue
        return tmpl

    return _DEFAULT_TEMPLATE


def _render_computed_field(name: str, ctx: dict) -> list[str]:
    """Renderiza un campo computado del frontmatter como lista de líneas YAML."""
    if name == "fecha":
        return [f"fecha: {ctx['fecha']}"]
    if name == "de":
        return [f"de: {ctx['sender_out']}"]
    if name == "para":
        lines = ["para:"]
        for addr in (ctx["to_list"] or [ctx["to_raw"]]):
            lines.append(f"  - {addr}")
        return lines
    if name == "cc":
        lines = ["cc:"]
        for addr in ctx["cc_list"]:
            lines.append(f"  - {addr}")
        return lines
    if name == "asunto":
        return [f"asunto: {ctx['subject_clean']}"]
    if name == "direccion":
        return [f"direccion: {ctx['direction']}"]
    if name == "adjuntos":
        lines = ["adjuntos:"]
        for n in ctx["att_names"]:
            lines.append(f"  - {n}")
        return lines
    return []


def _yaml_scalar(value: str) -> str:
    """Envuelve en comillas dobles strings que contienen caracteres especiales YAML."""
    if any(c in value for c in ('[[', ']]', ':', '#', '*', '&', '!', '{', '}')):
        return f'"{value.replace(chr(34), chr(92) + chr(34))}"'
    return value


def _render_fm_field(field, ctx: dict) -> list[str]:
    """Renderiza un ítem del frontmatter del template (string o dict) como líneas YAML."""
    if isinstance(field, str):
        return _render_computed_field(field, ctx)
    # Fixed: {"name": "tipo", "value": "correo"} o {"name": "tags", "value": ["correo"]}
    name  = field.get("name", "")
    value = field.get("value")
    if isinstance(value, list):
        lines = [f"{name}:"]
        for v in value:
            lines.append(f"  - {_yaml_scalar(str(v)) if isinstance(v, str) else v}")
        return lines
    elif value is None:
        return [f"{name}:"]
    else:
        return [f"{name}: {_yaml_scalar(str(value))}"]


# ─── Construcción de Markdown ─────────────────────────────────────────────────

def _build_md(subject: str, sender: str, to: str, cc: str,
              date_raw_dt, body: str, attachments: list[str],
              index: int = None, total: int = None,
              template: dict = None) -> str:
    """Construye el Markdown con YAML frontmatter según el template activo."""
    from email.utils import parsedate_to_datetime

    if template is None:
        template = _find_template(sender, to, cc, subject)

    # Fecha en yyyy-MM-dd
    try:
        if isinstance(date_raw_dt, str):
            fecha = parsedate_to_datetime(date_raw_dt).strftime("%Y-%m-%d")
        else:
            fecha = date_raw_dt.strftime("%Y-%m-%d")
    except Exception:
        fecha = str(date_raw_dt)

    aliases    = _load_aliases()
    to_list    = [_apply_alias(a, aliases) for a in _parse_addresses(to)]
    cc_list    = [_apply_alias(a, aliases) for a in _parse_addresses(cc)]
    att_names  = _attachment_names(attachments)
    sender_out = _apply_alias(sender, aliases)
    direction  = _detect_direction(sender)

    subject_clean = re.sub(r'[<>:";/\\|?*\x00-\x1f\[\]]', ' ', subject)
    subject_clean = re.sub(r'[,;]+', ' ', subject_clean)
    subject_clean = re.sub(r' {2,}', ' ', subject_clean).strip()

    title = subject
    if index is not None and total is not None and total > 1:
        title = f"{subject} [{index}/{total}]"

    ctx = {
        "fecha":         fecha,
        "sender_out":    sender_out,
        "to_list":       to_list,
        "to_raw":        _apply_alias(to, aliases),
        "cc_list":       cc_list,
        "subject_clean": subject_clean,
        "direction":     direction,
        "att_names":     att_names,
    }

    # ── YAML frontmatter ──────────────────────────────────────────────
    fm = ["---"]
    for field in template.get("frontmatter", _DEFAULT_TEMPLATE["frontmatter"]):
        fm.extend(_render_fm_field(field, ctx))
    fm.append("---")

    # ── Cuerpo ────────────────────────────────────────────────────────
    placeholders = {"fecha": ctx["fecha"], "subject": ctx["subject_clean"], "de": ctx["sender_out"]}

    body_header = template.get("body_header") or "## Contenido"
    for k, v in placeholders.items():
        body_header = body_header.replace(f"{{{k}}}", v)

    body_footer = template.get("body_footer", "")
    for k, v in placeholders.items():
        body_footer = body_footer.replace(f"{{{k}}}", v)

    if template.get("body_quote"):
        lines = body.strip().splitlines() if body else []
        body_text = "\n".join(f"> {line}" if line.strip() else ">" for line in lines) or "> *(Sin contenido)*"
    else:
        body_text = body.strip() if body else "*(Sin contenido)*"

    body_md = ["", f"# {title}", "", body_header, "", body_text]
    if body_footer:
        body_md.extend(["", body_footer])

    return "\n".join(fm + body_md)


def _seg_stem(seg_date, fallback_date_raw, subject_slug: str,
              filename_format: str = "{date} — {subject}") -> str:
    """Genera el stem del nombre de archivo usando la fecha del segmento o fallback."""
    from email.utils import parsedate_to_datetime
    from datetime import datetime

    if seg_date and isinstance(seg_date, datetime):
        date_str = seg_date.strftime('%Y-%m-%d-%H%M')
    else:
        try:
            if isinstance(fallback_date_raw, datetime):
                date_str = fallback_date_raw.strftime('%Y-%m-%d-%H%M')
            else:
                date_str = parsedate_to_datetime(fallback_date_raw).strftime('%Y-%m-%d-%H%M')
        except Exception:
            date_str = "0000-00-00-0000"

    return filename_format.replace("{date}", date_str).replace("{subject}", subject_slug)
