from pathlib import Path


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
