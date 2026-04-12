"""
Conversión de formatos adicionales via markitdown (Microsoft).

Formatos: .epub, .json, .xml, .zip
Instalación: pip install 'markitdown[all]'
"""

from pathlib import Path


MARKITDOWN_EXTENSIONS = {".epub", ".json", ".xml", ".zip"}


def convert_with_markitdown(path: Path) -> str:
    """Convierte un archivo a Markdown usando markitdown."""
    from markitdown import MarkItDown

    md = MarkItDown(enable_plugins=False)
    result = md.convert(str(path))
    return result.text_content or ""
