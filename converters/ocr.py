from pathlib import Path
import re

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def convert_image(path: Path) -> str:
    """OCR de imagen a Markdown. Detecta texto y tablas (requiere Tesseract)."""
    try:
        import pytesseract
    except ImportError:
        raise ImportError(
            "pytesseract no está instalado. Ejecuta: pip install pytesseract\n"
            "También necesitas instalar Tesseract OCR:\n"
            "  Windows: winget install UB-Mannheim.TesseractOCR\n"
            "  Descarga: https://github.com/UB-Mannheim/tesseract/wiki"
        )
    try:
        from PIL import Image
    except ImportError:
        raise ImportError("Pillow no está instalado. Ejecuta: pip install Pillow")

    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError(
            "Tesseract OCR no está instalado.\n"
            "Instálalo con:  winget install UB-Mannheim.TesseractOCR\n"
            "Descarga:  https://github.com/UB-Mannheim/tesseract/wiki\n"
            "Después reinicia el servidor."
        )

    img = Image.open(path).convert("RGB")
    lang = _detect_available_lang(pytesseract)

    # ── Extracción de tablas con img2table (opcional) ──────────────────────────
    tables_by_y: list[tuple[int, str]] = []
    table_bboxes: list[tuple[int, int, int, int]] = []

    try:
        from img2table.document import Image as Img2Doc
        from img2table.ocr import TesseractOCR

        print("  🔍 Detectando tablas...")
        ocr_engine = TesseractOCR(lang=lang)
        doc = Img2Doc(src=str(path))
        extracted = doc.extract_tables(
            ocr=ocr_engine,
            implicit_rows=False,
            borderless_tables=True,
            min_confidence=50,
        )
        page_tables = extracted.get(0, []) if isinstance(extracted, dict) else (extracted or [])

        for table in page_tables:
            if table is None:
                continue
            df = getattr(table, "df", None)
            if df is None or df.empty:
                continue
            bbox = table.bbox
            table_bboxes.append((bbox.x1, bbox.y1, bbox.x2, bbox.y2))
            md = _df_to_md(df)
            if md:
                tables_by_y.append((bbox.y1, md))
                print(f"  📊 Tabla detectada (y={bbox.y1})")

    except ImportError:
        pass  # img2table no disponible → solo texto

    # ── OCR de texto (enmascarando regiones de tabla) ─────────────────────────
    text_img = img.copy()
    if table_bboxes:
        from PIL import ImageDraw
        draw = ImageDraw.Draw(text_img)
        for x1, y1, x2, y2 in table_bboxes:
            draw.rectangle([x1, y1, x2, y2], fill=(255, 255, 255))

    print("  📝 Extrayendo texto...")
    raw = pytesseract.image_to_string(text_img, lang=lang, config="--psm 3")
    text = _clean_text(raw)

    # ── Combinar ──────────────────────────────────────────────────────────────
    parts = []
    if text:
        parts.append(text)
    if tables_by_y:
        if text:
            parts.append("---")
        for _, tmd in sorted(tables_by_y):
            parts.append(tmd)

    return "\n\n".join(parts)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _detect_available_lang(pytesseract) -> str:
    """Usa spa+eng si el pack de español está instalado, si no eng."""
    try:
        langs = pytesseract.get_languages()
        if "spa" in langs:
            return "spa+eng"
    except Exception:
        pass
    return "eng"


def _df_to_md(df) -> str:
    """Convierte DataFrame de img2table a tabla Markdown."""
    headers = [str(c) if c is not None else "" for c in df.columns]
    if not any(headers):
        return ""
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for _, row in df.iterrows():
        lines.append("| " + " | ".join(str(v) if v is not None else "" for v in row) + " |")
    return "\n".join(lines)


def _clean_text(text: str) -> str:
    """Elimina artefactos comunes del OCR."""
    lines = []
    for line in text.split("\n"):
        line = line.rstrip()
        if re.match(r'^[\s\-_=|~`]{3,}$', line):
            continue
        lines.append(line)
    result = re.sub(r'\n{3,}', '\n\n', "\n".join(lines))
    return result.strip()
