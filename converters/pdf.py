from pathlib import Path


def convert_pdf(path: Path) -> str:
    """Extrae texto de PDF con pdfplumber y lo estructura como Markdown."""
    import pdfplumber

    lines = []
    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages, 1):
            print(f"  📄 Procesando página {i}/{total}...", end="\r")

            text = page.extract_text()
            if text:
                lines.append(text.strip())

            for table in page.extract_tables():
                if not table:
                    continue
                header = table[0]
                rows = table[1:]
                lines.append("\n| " + " | ".join(str(c or "") for c in header) + " |")
                lines.append("| " + " | ".join("---" for _ in header) + " |")
                for row in rows:
                    lines.append("| " + " | ".join(str(c or "") for c in row) + " |")
                lines.append("")

    print()
    return "\n\n".join(lines)
