# MD Converter

Conversor universal de archivos a Markdown, con UI web local y soporte especializado para correos de Outlook.

---

## Archivos del proyecto

```
mdconverter/
├── convert_to_md.py           # CLI + dispatcher (re-exporta converters/)
├── converter_ui.py            # Servidor Flask con UI drag & drop (localhost:5000)
├── contact_aliases.json       # Reglas de alias para contactos conocidos
├── converters/                # Lógica de conversión por formato
│   ├── docx.py                #   .docx → Markdown (mammoth)
│   ├── pdf.py                 #   .pdf  → Markdown (pdfplumber)
│   ├── html.py                #   HTML/URL → Markdown + soporte de tablas HTML
│   ├── tabular.py             #   .xlsx / .csv → tabla Markdown (pandas)
│   └── email/
│       ├── thread.py          #   Detección y separación de hilos
│       ├── builders.py        #   Construcción de frontmatter YAML y alias
│       ├── eml.py             #   Conversor .eml
│       └── msg.py             #   Conversor .msg (Outlook)
├── md_output/                 # Salida de archivos .md (se crea automáticamente)
└── correos/                   # Carpeta vigilada por defecto para .msg de Outlook
```

---

## Instalación

**Requisitos:** Python 3.10 o superior.

```powershell
pip install flask mammoth pdfplumber html2text pandas openpyxl tabulate requests beautifulsoup4 extract-msg
```

---

## Uso

### UI web (recomendado)

```powershell
python converter_ui.py
```

Se abre automáticamente el browser en `http://localhost:5000`. Para lanzarlo sin ventana de terminal, usa el archivo `iniciar.bat`:

```bat
@echo off
cd /d %~dp0
python converter_ui.py
pause
```

### CLI

```powershell
# Un archivo
python convert_to_md.py archivo.docx
python convert_to_md.py reporte.pdf
python convert_to_md.py correo.msg
python convert_to_md.py datos.xlsx
python convert_to_md.py https://ejemplo.com

# Múltiples archivos
python convert_to_md.py *.msg *.pdf

# Carpeta completa (recursivo)
python convert_to_md.py ./documentos/ -o ./markdown_output/
```

---

## Formatos soportados

| Formato | Extensión | Notas |
|---|---|---|
| Word | `.docx` | Preserva encabezados, listas, negritas y links |
| PDF | `.pdf` | Extrae texto y tablas por página |
| Excel | `.xlsx` | Cada hoja genera una sección con tabla Markdown |
| CSV | `.csv` | Tabla Markdown directa |
| HTML / Web | `.html`, `.htm`, URL | Mantiene links e imágenes como referencias |
| Email Outlook | `.msg` | Ver sección de correos abajo |
| Email estándar | `.eml` | Ver sección de correos abajo |

---

## Conversión de correos (.msg / .eml)

### Nombre de archivo generado

```
yyyy-MM-dd — {Asunto}.md
```

Ejemplo: `2026-04-03 — Re Prueba piloto SMS+.md`

### Conversaciones (hilos)

Cuando el correo contiene una conversación de ida y vuelta, **se genera un archivo `.md` separado por cada mensaje** del hilo:

```
2026-04-03 — Re Prueba piloto SMS+ — msg01 de 18.md
2026-04-02 — Re Prueba piloto SMS+ — msg02 de 18.md
...
2026-02-23 — Prueba piloto SMS+ — msg18 de 18.md
```

Cada archivo tiene la **fecha real** del mensaje correspondiente (no la del correo más reciente), y los metadatos (`de`, `para`, `cc`, `asunto`) son los del mensaje específico cuando están disponibles.

El separador de hilo se detecta automáticamente en cinco formatos:

- `On [fecha] [nombre] wrote:` — estilo Gmail / Outlook Web
- `________________________________` seguido de bloque `De:/Enviado:` — estilo Outlook Desktop
- Bloque `De:/From:` standalone (sin separador previo) — formato mixto
- `**From:**` / `**Sent:**` en negritas — generado por html2text desde el HTML del correo
- `--- Original Message ---` — clientes de correo alternativos

Las tablas HTML presentes en el cuerpo del correo se convierten a tablas Markdown.

### Frontmatter YAML

Cada `.md` de correo incluye un frontmatter con metadatos estructurados:

```yaml
---
fecha: 2026-04-03
de: Stefan Cirlan <stefan.cirlan@voxsolutions.co>
para:
  - Clara Andrea Cabrera Delgado <Clara.Cabrera@nuevatel.com>
cc:
  - Tudor Marasescu <tudor.marasescu@voxsolutions.co>
  - Crenguta Craciun <crenguta.craciun@voxsolutions.co>
asunto: Re Prueba piloto SMS+
tipo: correo
direccion: enviado
tags:
  - correo
adjuntos:
  - reporte_q1.pdf
---

# Re Prueba piloto SMS+

## Contenido

Hola Clara, ...
```

El campo `direccion` se determina automáticamente: es `enviado` si el remitente es `Alonzo.Vera` o `alvera`, y `recibido` en cualquier otro caso.

---

## UI web — Funcionalidades

### Drag & Drop
Arrastra archivos directamente desde el Explorador de Windows a la zona de drop del browser. Soporta todos los formatos listados arriba.

### Carpeta vigilada (para Outlook)
Outlook no puede arrastrar correos directamente al browser. El flujo recomendado es:

1. En Outlook: `Archivo → Guardar como` → seleccionar la carpeta vigilada → guardar como `.msg`
2. La UI detecta el archivo nuevo automáticamente y lo convierte en segundos

La carpeta vigilada por defecto es `D:\mdconverter\correos\` y se inicia automáticamente al arrancar el servidor. Se puede cambiar desde el panel derecho de la UI.

### Botones de acceso rápido
- **📂 Correos** — abre la carpeta de entrada en el Explorador
- **📄 Salida MD** — abre la carpeta `md_output\` donde se guardan los archivos generados

---

## Carpetas por defecto

| Carpeta | Ruta | Descripción |
|---|---|---|
| Correos (entrada) | `D:\mdconverter\correos\` | Carpeta vigilada para `.msg` de Outlook |
| Markdown (salida) | `D:\mdconverter\md_output\` | Archivos `.md` generados |

Ambas se crean automáticamente si no existen.

---

## Aliases de contactos

Los campos `de`, `para` y `cc` de los correos se normalizan usando reglas definidas en `contact_aliases.json`. Cuando una dirección coincide con algún fragmento de `match`, se reemplaza por el valor de `alias`.

```json
{
  "aliases": [
    {
      "alias": "\"[[Clara Cabrera]]\"",
      "match": ["Clara Cabrera", "Clara.Cabrera"]
    },
    {
      "alias": "\"[[Alonzo Vera]]\"",
      "match": ["Alonzo Vera", "Alonzo.Vera", "alvera"]
    },
    {
      "alias": "\"[[José Revollo]]\"",
      "match": ["José Revollo", "Jose Revollo", "José.Revollo", "Jose.Revollo"]
    },
    {
      "alias": "\"[[Liuba Sarmiento]]\"",
      "match": ["Liuba Sarmiento", "Liuba.Sarmiento"]
    }
  ]
}
```

El resultado en el frontmatter:

```yaml
de: Stefan Cirlan <stefan.cirlan@voxsolutions.co>
para:
  - "[[Clara Cabrera]]"
  - Tudor Marasescu <tudor.marasescu@voxsolutions.co>
cc:
  - "[[Alonzo Vera]]"
  - "[[José Revollo]]"
```

**Para agregar un contacto:**

1. Abrir `contact_aliases.json`
2. Añadir un bloque al array `aliases`:

```json
{
  "alias": "\"[[Nombre Apellido]]\"",
  "match": ["Nombre Apellido", "nombre.apellido"]
}
```

Los cambios tienen efecto inmediato — no es necesario reiniciar el servidor.

---

## Dependencias

| Librería | Uso |
|---|---|
| `flask` | Servidor web de la UI |
| `mammoth` | Conversión `.docx → md` |
| `pdfplumber` | Extracción de texto y tablas de PDF |
| `html2text` | Conversión HTML → Markdown |
| `pandas` + `openpyxl` + `tabulate` | Conversión `.xlsx` / `.csv` |
| `requests` + `beautifulsoup4` | Descarga y parseo de URLs |
| `extract-msg` | Lectura de archivos `.msg` de Outlook |
