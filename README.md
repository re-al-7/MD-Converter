# MD Converter

Conversor universal de archivos a Markdown, con UI web local y soporte especializado para correos de Outlook.

---

## Archivos del proyecto

```
mdconverter/
├── convert_to_md.py           # CLI + dispatcher (re-exporta converters/)
├── converter_ui.py            # Servidor Flask con UI drag & drop (localhost:3200)
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
pip install flask mammoth pdfplumber html2text pandas openpyxl tabulate requests beautifulsoup4 extract-msg pytesseract Pillow img2table
```

> **OCR de imágenes:** además del paquete `pytesseract`, necesitas instalar el binario de Tesseract:
> ```powershell
> winget install UB-Mannheim.TesseractOCR
> ```
> Para reconocimiento en español, selecciona el pack de idioma **Spanish** durante la instalación (o descárgalo desde la misma página). `img2table` es opcional pero mejora significativamente la detección de tablas.

---

## Uso

### UI web (recomendado)

```powershell
python converter_ui.py
```

Se abre automáticamente el browser en `http://localhost:3200`. Para lanzarlo sin ventana de terminal, usa el archivo `iniciar.bat`:

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
| Imagen | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.webp` | OCR con Tesseract — ver sección abajo |

---

## Conversión de imágenes (OCR)

Convierte imágenes a Markdown extrayendo texto y tablas mediante OCR.

### Formatos soportados

`.jpg` · `.jpeg` · `.png` · `.bmp` · `.tiff` · `.tif` · `.webp`

### Cómo funciona

1. **Detección de tablas** — si `img2table` está instalado, detecta automáticamente regiones de tabla y las convierte a tablas Markdown. Las regiones de tabla se excluyen del paso siguiente.
2. **Extracción de texto** — `pytesseract` aplica OCR al resto de la imagen. Usa español + inglés si el pack de idioma `spa` está disponible.
3. **Salida** — texto primero, luego tablas separadas por `---`.

### Ejemplo de salida

```markdown
Informe de ventas Q1 2026

Los resultados del trimestre superaron las expectativas en un 12%...

---

| Región | Ventas | Variación |
| --- | --- | --- |
| Norte | 1.200.000 | +8% |
| Sur | 980.000 | +15% |
```

### Requisitos

| Componente | Instalación |
|---|---|
| `pytesseract` | `pip install pytesseract` |
| `Pillow` | `pip install Pillow` |
| Tesseract OCR (binario) | `winget install UB-Mannheim.TesseractOCR` |
| `img2table` (opcional, tablas) | `pip install img2table` |

> Si Tesseract no está en el PATH de Windows, agrégalo manualmente: `C:\Program Files\Tesseract-OCR`.

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

## Inicio automático con Windows

Para que la aplicación arranque sola cada vez que inicies sesión en Windows, el proyecto incluye dos archivos:

| Archivo | Descripción |
|---|---|
| `run_hidden.vbs` | Lanza `converter_ui.py` sin mostrar ventana de consola |
| `setup_startup.py` | Instala o desinstala la entrada en el registro de Windows |

### Activar el inicio automático

```powershell
python setup_startup.py install
```

Esto agrega una entrada en `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (sin necesidad de permisos de administrador). A partir del próximo inicio de sesión, la app arrancará en segundo plano y abrirá `http://localhost:3200` en el navegador.

### Otros comandos

```powershell
# Verificar si está instalado
python setup_startup.py status

# Desactivar el inicio automático
python setup_startup.py uninstall
```

> **Nota:** Si usas un entorno virtual (`venv/`), el script lo detecta automáticamente. En caso contrario, usa el `pythonw.exe` del sistema.

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
| `pytesseract` | Wrapper de Python para Tesseract OCR |
| `Pillow` | Carga y preprocesamiento de imágenes |
| `img2table` *(opcional)* | Detección de tablas en imágenes |
