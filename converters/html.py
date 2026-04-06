import re
from pathlib import Path


def _html_table_to_md(table) -> str:
    """
    Convierte un elemento BeautifulSoup <table> a tabla Markdown.
    Solo procesa filas directas (no tablas anidadas).
    """
    rows_data = []
    containers = [table] + table.find_all(['thead', 'tbody', 'tfoot'], recursive=False)
    for container in containers:
        for tr in container.find_all('tr', recursive=False):
            cells = []
            for cell in tr.find_all(['th', 'td'], recursive=False):
                text = cell.get_text(separator=' ', strip=True)
                text = text.replace('|', '\\|')
                text = re.sub(r'\s+', ' ', text).strip()
                cells.append(text)
            if cells:
                rows_data.append(cells)

    if not rows_data:
        return ''

    ncols = max(len(r) for r in rows_data)
    padded = [r + [''] * (ncols - len(r)) for r in rows_data]

    header = padded[0]
    lines = [
        '| ' + ' | '.join(header) + ' |',
        '| ' + ' | '.join('---' for _ in header) + ' |',
    ]
    for row in padded[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')

    return '\n'.join(lines)


def _html_to_md_with_tables(html: str) -> str:
    """
    Convierte HTML a Markdown convirtiendo las tablas HTML en tablas Markdown
    reales antes de pasar el resto a html2text.
    """
    import html2text
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, 'html.parser')
    tables_md: dict[str, str] = {}

    for i, table in enumerate(soup.find_all('table')):
        md = _html_table_to_md(table)
        if md:
            placeholder = f'___TABLE_{i}___'
            tables_md[placeholder] = md
            new_tag = soup.new_tag('div')
            new_tag.string = f'\n{placeholder}\n'
            table.replace_with(new_tag)

    converter = html2text.HTML2Text()
    converter.ignore_links  = False
    converter.body_width    = 0
    converter.unicode_snob  = True
    converter.bypass_tables = True

    result = converter.handle(str(soup))

    for placeholder, md in tables_md.items():
        result = result.replace(placeholder, f'\n{md}\n')

    return result.strip()


def convert_html(source: str, is_url: bool = False) -> str:
    """Convierte HTML local o URL a Markdown."""
    import html2text
    import requests

    converter = html2text.HTML2Text()
    converter.ignore_links  = False
    converter.ignore_images = False
    converter.body_width    = 0
    converter.protect_links = True
    converter.unicode_snob  = True

    if is_url:
        print(f"  🌐 Descargando {source}...")
        resp = requests.get(source, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        html_content = resp.text
    else:
        with open(source, "r", encoding="utf-8", errors="replace") as f:
            html_content = f.read()

    return converter.handle(html_content)
