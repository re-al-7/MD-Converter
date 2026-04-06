from pathlib import Path


def convert_docx(path: Path) -> str:
    """Convierte .docx a Markdown usando mammoth."""
    import mammoth

    with open(path, "rb") as f:
        result = mammoth.convert_to_markdown(f)

    if result.messages:
        warnings = [m.message for m in result.messages]
        print(f"  ⚠️  Advertencias: {'; '.join(warnings)}")

    return result.value
