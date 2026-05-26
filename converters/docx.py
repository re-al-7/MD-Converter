import re
from pathlib import Path

from .html import _html_to_md_with_tables


def convert_docx(path: Path) -> str:
    """Convierte .docx a Markdown usando mammoth (vía HTML para preservar tablas)."""
    import mammoth

    with open(path, "rb") as f:
        result = mammoth.convert_to_html(f)

    if result.messages:
        warnings = [m.message for m in result.messages]
        print(f"  ⚠️  Advertencias: {'; '.join(warnings)}")

    md = _html_to_md_with_tables(result.value)
    # html2text escapa el punto en headings numerados ("# 1\. Título") de forma
    # innecesaria — el contexto de heading no puede confundirse con lista ordenada.
    md = re.sub(r'^(#{1,6} \d+)\\\.', r'\1.', md, flags=re.MULTILINE)
    return md
