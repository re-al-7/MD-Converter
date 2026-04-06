#!/usr/bin/env python3
"""
convert_to_md.py — Conversor universal a Markdown
Soporta: .docx, .pdf, .pptx, .html, .htm, .xlsx, .csv, .eml, .msg, URLs web

Uso:
    python convert_to_md.py archivo.docx
    python convert_to_md.py presentacion.pptx
    python convert_to_md.py reporte.pdf
    python convert_to_md.py datos.xlsx
    python convert_to_md.py pagina.html
    python convert_to_md.py correo.eml
    python convert_to_md.py https://ejemplo.com
    python convert_to_md.py carpeta/           # Convierte todos los archivos soportados
"""

import sys
import argparse
from pathlib import Path

# Re-exportar para compatibilidad con converter_ui.py y otros importadores
from converters import (
    convert_docx,
    convert_pdf,
    convert_html,
    convert_xlsx,
    convert_csv,
    convert_pptx,
    convert_eml,
    convert_msg,
)

SUPPORTED_EXTENSIONS = {".docx", ".pdf", ".pptx", ".html", ".htm", ".xlsx", ".csv", ".eml", ".msg"}


def install_deps():
    """Instala dependencias si no están presentes."""
    import subprocess
    deps = [
        "mammoth",        # docx → md
        "pdfplumber",     # pdf → texto
        "python-pptx",    # pptx → md
        "html2text",      # html → md
        "pandas",         # xlsx/csv → md
        "openpyxl",       # motor Excel para pandas
        "tabulate",       # formato tabla markdown
        "requests",       # fetch URL
        "beautifulsoup4", # parseo HTML
        "extract-msg",    # .msg Outlook → datos estructurados
    ]
    print("📦 Instalando dependencias...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + deps)
    print("✅ Dependencias listas.\n")


def convert_file(source: str, output_dir: Path = None) -> Path | None:
    """Detecta el tipo de archivo y lo convierte. Devuelve la ruta del .md generado."""
    is_url = source.startswith("http://") or source.startswith("https://")

    if is_url:
        stem = source.split("//")[-1].split("/")[0].replace(".", "_")
        ext  = ".html"
    else:
        path = Path(source)
        if not path.exists():
            print(f"❌ No encontrado: {source}")
            return None
        stem = path.stem
        ext  = path.suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            print(f"⚠️  Formato no soportado: {ext} ({path.name})")
            return None

    out_dir  = output_dir or (Path(source).parent if not is_url else Path("."))
    out_path = out_dir / f"{stem}.md"

    print(f"\n🔄 Convirtiendo: {source}")
    print(f"   → {out_path}")

    try:
        if is_url or ext in (".html", ".htm"):
            content = convert_html(source, is_url=is_url)
        elif ext == ".docx":
            content = convert_docx(Path(source))
        elif ext == ".pptx":
            content = convert_pptx(Path(source))
        elif ext == ".pdf":
            content = convert_pdf(Path(source))
        elif ext == ".xlsx":
            content = convert_xlsx(Path(source))
        elif ext == ".csv":
            content = convert_csv(Path(source))
        elif ext in (".eml", ".msg"):
            fn           = convert_eml if ext == ".eml" else convert_msg
            mail_results = fn(Path(source))
            saved        = []
            for fname, md in mail_results:
                fpath = out_dir / fname
                fpath.parent.mkdir(parents=True, exist_ok=True)
                fpath.write_text(md, encoding="utf-8")
                print(f"   ✅ {fname} ({fpath.stat().st_size / 1024:.1f} KB)")
                saved.append(fpath)
            return saved[0] if saved else None
        else:
            print(f"❌ Sin conversor para {ext}")
            return None

        out_dir.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
        print(f"   ✅ Guardado ({out_path.stat().st_size / 1024:.1f} KB)")
        return out_path

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return None


def convert_folder(folder: Path, output_dir: Path = None) -> list[Path]:
    """Convierte todos los archivos soportados en una carpeta."""
    files   = [f for f in folder.rglob("*") if f.suffix.lower() in SUPPORTED_EXTENSIONS]
    results = []

    if not files:
        print(f"⚠️  No se encontraron archivos soportados en {folder}")
        return results

    print(f"📂 Encontrados {len(files)} archivos en {folder}")
    for f in files:
        result = convert_file(str(f), output_dir)
        if result:
            results.append(result)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Conversor universal a Markdown (.docx, .pdf, .html, .xlsx, .csv, .eml, .msg, URLs)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("sources", nargs="+", metavar="ARCHIVO_O_URL",
                        help="Archivos, carpetas o URLs a convertir")
    parser.add_argument("-o", "--output", metavar="DIR",
                        help="Carpeta destino para los .md generados")
    parser.add_argument("--install", action="store_true",
                        help="Instalar dependencias antes de convertir")

    args = parser.parse_args()

    if args.install:
        install_deps()

    output_dir = Path(args.output) if args.output else None
    converted  = []

    for source in args.sources:
        path = Path(source)
        if not (source.startswith("http://") or source.startswith("https://")) and path.is_dir():
            converted.extend(convert_folder(path, output_dir))
        else:
            result = convert_file(source, output_dir)
            if result:
                converted.append(result)

    print(f"\n{'─' * 40}")
    print(f"✨ Convertidos: {len(converted)} archivo(s)")
    for f in converted:
        print(f"   • {f}")


if __name__ == "__main__":
    main()
