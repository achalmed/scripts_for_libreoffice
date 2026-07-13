# pdf-suite v3.0

> Suite completa de manipulación PDF para Linux — el equivalente a iLovePDF desde tu terminal.
> Comprimir · Unir · Dividir · Extraer · Rotar · Convertir · OCR · Metadatos · Proteger · Reparar.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](<https://img.shields.io/badge/platform-Linux%20(Kubuntu%20%7C%20Arch)-informational.svg>)
![Shell](https://img.shields.io/badge/shell-bash%205.x-89e051.svg)
![Version](https://img.shields.io/badge/version-3.0.0-green.svg)

**Autor:** Edison Achalma · UNSCH · Ayacucho, Perú  
**GitHub:** [@achalmed](https://github.com/achalmed)

---

## Tabla de Contenidos

- [Descripción](#descripción)
- [Qué incluye](#qué-incluye)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Uso](#uso)
- [Operaciones detalladas](#operaciones-detalladas)
- [Arquitectura](#arquitectura)
- [Bugs corregidos](#bugs-corregidos)
- [Solución de problemas](#solución-de-problemas)
- [Cómo contribuir](#cómo-contribuir)
- [Notas y advertencias](#notas-y-advertencias)

---

## Descripción

`pdf-suite` unifica tres scripts previos en un único proyecto modular:

| Script original                 | Estado en v3.0                  |
| ------------------------------- | ------------------------------- |
| `compress_pdf.sh` v2.0          | Integrado en `lib/compress.sh`  |
| `arch_pdf_metadata_commands.sh` | Integrado en `lib/metadata.sh`  |
| `test_compression.sh`           | Integrado como operación `test` |

Además agrega operaciones completamente nuevas: merge, split, extract, rotate, reorder, delete, convert, OCR, protect, watermark, repair, validate, optimize.

---

## Qué incluye

| Operación   | Descripción                                           | Motores                      |
| ----------- | ----------------------------------------------------- | ---------------------------- |
| `compress`  | Comprimir PDFs (5 niveles de calidad)                 | Ghostscript, ocrmypdf        |
| `merge`     | Unir múltiples PDFs en uno                            | qpdf                         |
| `split`     | Dividir un PDF en grupos de N páginas                 | qpdf                         |
| `extract`   | Extraer páginas específicas (rangos, no consecutivas) | qpdf                         |
| `rotate`    | Rotar páginas (90°, 180°, 270°, rangos)               | qpdf                         |
| `reorder`   | Reordenar páginas (ej: invertir)                      | qpdf                         |
| `delete`    | Eliminar páginas de un PDF                            | qpdf + python3               |
| `convert`   | PDF → PNG/JPG/SVG/TXT/HTML; imágenes → PDF            | pdftoppm, pdftotext, img2pdf |
| `ocr`       | OCR a PDFs escaneados (añade capa de texto buscable)  | ocrmypdf, tesseract          |
| `metadata`  | Ver/editar metadatos XMP y DocInfo                    | exiftool, pdfinfo            |
| `protect`   | Cifrar (AES-256) o descifrar PDFs                     | qpdf                         |
| `watermark` | Marca de agua de texto o sello PDF                    | cpdf, pdftk                  |
| `repair`    | Reparar PDFs corruptos (3 estrategias en cascada)     | qpdf, mutool, ghostscript    |
| `validate`  | Validar estructura interna del PDF                    | qpdf, pdfinfo, mutool        |
| `optimize`  | Linearizar + comprimir streams (Fast Web View)        | qpdf                         |
| `info`      | Información completa: páginas, tamaño, metadatos      | pdfinfo, exiftool            |
| `test`      | Comparar todos los métodos de compresión en un PDF    | Ghostscript, ocrmypdf        |
| `deps`      | Verificar estado de todas las dependencias            | —                            |

---

## Requisitos

### Sistema operativo

- Ubuntu 22.04+ / Kubuntu 22.04+ (probado)
- Arch Linux / Archcraft (compatible con ajuste del instalador)
- Cualquier Linux con Bash ≥ 5.0

### Dependencias obligatorias (sin estas el suite no arranca)

| Herramienta | Paquete       | Para qué                    |
| ----------- | ------------- | --------------------------- |
| `gs`        | ghostscript   | Compresión y conversión PDF |
| `qpdf`      | qpdf          | Manipulación estructural    |
| `pdfinfo`   | poppler-utils | Información básica de PDFs  |

### Dependencias opcionales (activan operaciones específicas)

| Herramienta  | Paquete                    | Activa                            |
| ------------ | -------------------------- | --------------------------------- |
| `pdftk`      | pdftk                      | Sellos PDF, formularios           |
| `pdftotext`  | poppler-utils              | Extracción de texto               |
| `pdfimages`  | poppler-utils              | Extracción de imágenes            |
| `pdftoppm`   | poppler-utils              | PDF → PNG/JPEG                    |
| `pdftohtml`  | poppler-utils              | PDF → HTML                        |
| `pdftocairo` | poppler-utils              | PDF → SVG                         |
| `mutool`     | mupdf-tools                | Reparación, info adicional        |
| `ocrmypdf`   | ocrmypdf                   | OCR en PDFs escaneados            |
| `tesseract`  | tesseract-ocr              | Motor OCR                         |
| `exiftool`   | libimage-exiftool-perl     | Metadatos XMP completos           |
| `img2pdf`    | img2pdf                    | Imágenes → PDF sin pérdida        |
| `pdfjam`     | texlive-extra-utils        | N-up, diseño de páginas           |
| `pdfcrop`    | texlive-extra-utils        | Recortar márgenes                 |
| `convert`    | imagemagick                | Conversión de imágenes (fallback) |
| `cpdf`       | [manual, ver §Instalación] | Marca de agua texto, numeración   |

---

## Instalación

### Opción A — Rápida (recomendada para Edison)

```bash
# Clonar o copiar el proyecto a su ubicación definitiva
mkdir -p ~/Documents/scripts_for_linux/pdf-suite
cd ~/Documents/scripts_for_linux/pdf-suite

# Ejecutar el instalador
chmod +x install.sh
./install.sh
```

El instalador:

1. Detecta tu gestor de paquetes (apt / pacman)
2. Verifica e instala dependencias obligatorias
3. Ofrece instalar cada dependencia opcional
4. Copia los archivos a `~/Documents/scripts_for_linux/pdf-suite/`
5. Crea el wrapper `/usr/local/bin/pdf-suite` para ejecutar desde cualquier lugar

### Opción B — Manual (Ubuntu/Kubuntu)

```bash
# Dependencias obligatorias + recomendadas
sudo apt update && sudo apt install -y \
  ghostscript qpdf poppler-utils \
  pdftk mupdf-tools ocrmypdf \
  tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng \
  libimage-exiftool-perl img2pdf imagemagick \
  texlive-extra-utils

# Permisos
chmod +x ~/Documents/scripts_for_linux/pdf-suite/main.sh
chmod +x ~/Documents/scripts_for_linux/pdf-suite/lib/*.sh

# Wrapper global
sudo tee /usr/local/bin/pdf-suite > /dev/null << 'EOF'
#!/usr/bin/env bash
exec "${HOME}/Documents/scripts_for_linux/pdf-suite/main.sh" "$@"
EOF
sudo chmod +x /usr/local/bin/pdf-suite
```

### Instalación de cpdf (marca de agua y numeración)

```bash
wget https://github.com/coherentgraphics/cpdf-binaries/raw/master/Linux-Intel/cpdf
chmod +x cpdf
sudo mv cpdf /usr/local/bin/
```

---

## Uso

### Sintaxis

```
pdf-suite <operación> [opciones-globales] [opciones-de-operación] <archivo(s)>
pdf-suite                     # Sin argumentos: abre menú interactivo
```

### Opciones globales

| Flag              | Descripción                                              |
| ----------------- | -------------------------------------------------------- |
| `-v, --verbose`   | Mostrar detalles técnicos (stderr de gs, ocrmypdf, etc.) |
| `-n, --dry-run`   | Simular sin escribir ningún archivo                      |
| `-f, --force`     | Sobreescribir archivos de salida existentes              |
| `-r, --recursive` | Procesar subdirectorios en operaciones de directorio     |
| `-o, --output`    | Ruta explícita del archivo de salida                     |
| `-s, --suffix`    | Sufijo para archivos de salida (default: `_out`)         |
| `--log-file`      | Guardar log en `~/.local/share/pdf-suite/logs/`          |
| `--version`       | Mostrar versión                                          |
| `-h, --help`      | Mostrar ayuda completa                                   |

---

## Operaciones detalladas

### `compress` — Comprimir PDFs

```bash
# Un archivo con método ebook (recomendado para lectura)
pdf-suite compress -m ebook libro.pdf

# Toda la carpeta recursivamente
pdf-suite compress -m ebook -r ~/Documents/biblioteca

# Solo comprimir si reduce más del 20%
pdf-suite compress -m ebook -r -t 20 ~/Documents/biblioteca

# Máxima compresión (para email/web)
pdf-suite compress -m screen informe.pdf
```

| Método     | DPI | Reducción típica | Uso                               |
| ---------- | --- | ---------------- | --------------------------------- |
| `screen`   | 72  | 80–95%           | Web, email, máxima compresión     |
| `ebook`    | 150 | 60–85% ⭐        | **Recomendado** — lectura digital |
| `printer`  | 300 | 40–70%           | Para imprimir                     |
| `prepress` | 300 | 20–50%           | Impresión profesional             |
| `ocr`      | —   | 50–80% ⭐        | PDFs escaneados                   |

### `merge` — Unir PDFs

```bash
pdf-suite merge doc1.pdf doc2.pdf doc3.pdf -o unido.pdf

# Sin -o: genera merged_out.pdf junto al primer archivo
pdf-suite merge capitulo1.pdf capitulo2.pdf capitulo3.pdf
```

### `split` — Dividir PDF

```bash
# Dividir en páginas individuales
pdf-suite split libro.pdf

# Dividir en partes de 10 páginas
pdf-suite split --pages 10 libro.pdf
```

### `extract` — Extraer páginas

```bash
# Páginas 1 a 5
pdf-suite extract --pages 1-5 informe.pdf

# Páginas no consecutivas
pdf-suite extract --pages 1,3,7,10-12 informe.pdf

# Desde la página 10 hasta el final (z = última página)
pdf-suite extract --pages 10-z tesis.pdf

# Con nombre de salida explícito
pdf-suite extract --pages 28 -o portada.pdf documento.pdf
```

### `rotate` — Rotar páginas

```bash
# Rotar todas las páginas 90°
pdf-suite rotate --angle 90 documento.pdf

# Rotar solo páginas 1-3 a 180°
pdf-suite rotate --angle 180 --pages 1-3 documento.pdf

# Rotar en sentido antihorario
pdf-suite rotate --angle -90 documento.pdf
```

### `reorder` — Reordenar páginas

```bash
# Invertir el orden de todas las páginas
pdf-suite reorder documento.pdf

# Rango personalizado de reordenamiento
pdf-suite reorder --range "3,1,2,5,4" documento.pdf
```

### `delete` — Eliminar páginas

```bash
# Eliminar páginas 3, 5 y 10 a 12
pdf-suite delete --pages "3,5,10-12" documento.pdf
```

### `convert` — Convertir

```bash
# PDF → imágenes PNG a 300 DPI
pdf-suite convert --to png --dpi 300 presentacion.pdf

# PDF → imágenes JPEG a 150 DPI
pdf-suite convert --to jpg presentacion.pdf

# PDF → texto plano
pdf-suite convert --to txt paper.pdf

# PDF → HTML
pdf-suite convert --to html paper.pdf

# Imágenes → PDF (sin pérdida con img2pdf)
pdf-suite convert --to pdf scan1.jpg scan2.jpg -o documento.pdf

# Extraer imágenes embebidas en un PDF
pdf-suite convert --extract-images documento.pdf
```

### `ocr` — OCR en PDFs escaneados

```bash
# OCR en español
pdf-suite ocr -l spa escaneo.pdf

# OCR en español e inglés
pdf-suite ocr -l spa+eng escaneo.pdf

# OCR en lote (carpeta completa)
pdf-suite ocr -r -l spa ~/Documents/doc_administrativos

# Ver qué PDFs necesitan OCR antes de procesarlos
pdf-suite ocr --scan ~/Documents/biblioteca
```

### `metadata` — Ver y editar metadatos

```bash
# Ver información completa de un PDF
pdf-suite info documento.pdf
pdf-suite metadata documento.pdf

# Ver metadatos en JSON
pdf-suite metadata --json documento.pdf

# Escribir metadatos
pdf-suite metadata --set-title "Análisis Económico de Ayacucho" \
                   --set-author "Edison Achalma" \
                   --set-subject "Economía, UNSCH" \
                   --set-keywords "economía, Ayacucho, estadística" \
                   documento.pdf

# Título desde nombre de archivo (lote)
pdf-suite metadata --from-filename -r ~/Documents/biblioteca

# Ver qué PDFs de Zotero no tienen metadatos
pdf-suite metadata --check-missing ~/Zotero/storage
```

### `protect` — Cifrar y descifrar

```bash
# Cifrar con contraseña owner (obligatoria)
pdf-suite protect --encrypt --owner-pass "mi_clave_segura" tesis.pdf

# Cifrar con contraseña de usuario y owner
pdf-suite encrypt --user-pass "leer" --owner-pass "admin" tesis.pdf

# Descifrar
pdf-suite decrypt --password "mi_clave_segura" tesis_protected.pdf
```

### `watermark` — Marca de agua

```bash
# Marca de agua de texto diagonal
pdf-suite watermark --text "BORRADOR" documento.pdf

# Personalizar opacidad, ángulo y color
pdf-suite watermark --text "CONFIDENCIAL" \
                    --opacity 0.2 \
                    --angle 30 \
                    --color "0 0 0.8" \
                    documento.pdf

# Sello PDF (logo, firma) encima del contenido
pdf-suite protect --stamp logo.pdf documento.pdf

# Añadir números de página
pdf-suite protect --page-numbers documento.pdf
```

### `repair` — Reparar y validar

```bash
# Reparar un PDF dañado (3 estrategias en cascada)
pdf-suite repair documento_roto.pdf

# Solo validar estructura sin modificar
pdf-suite validate documento.pdf

# Optimizar para web (Fast Web View + compresión de streams)
pdf-suite optimize documento.pdf
```

### `test` — Comparar métodos de compresión

```bash
pdf-suite test libro_grande.pdf
```

Produce una tabla con el tamaño y reducción para cada método:

```
  Método       Tamaño         Reducción    Recomendación
  ──────────────────────────────────────────────────────
  screen       32.1 MB            79%      Web/email
  ebook        67.3 MB            56%      Lectura digital ⭐
  printer      89.4 MB            42%      Imprimir
  prepress    112.0 MB            27%      Impresión pro
  ocr          58.7 MB            62%      Escaneados ⭐
```

### `deps` — Verificar dependencias

```bash
pdf-suite deps
```

---

## Arquitectura

```
~/Documents/scripts_for_linux/pdf-suite/
├── main.sh            # Orquestador: sourcea módulos, despacha operaciones
├── config.sh          # Colores, rutas, defaults globales, PDF_SEARCH_PATHS
├── install.sh         # Instalador multiplataforma (apt / pacman)
├── README.md          # Esta documentación
└── lib/
    ├── logger.sh      # log_info/warn/error/debug/step/success/failure
    ├── validator.sh   # Validación de deps, archivos PDF, rangos, tamaños
    ├── cli.sh         # Parsing de flags globales, menú TUI, prompts
    ├── compress.sh    # compress_single_pdf, compress_batch, test_all_methods
    ├── manipulate.sh  # merge, split, extract, rotate, reorder, delete
    ├── convert.sh     # pdf→img, pdf→txt, pdf→html, img→pdf, extract images
    ├── metadata.sh    # show_pdf_info, read/write metadata, check_missing
    ├── ocr.sh         # apply_ocr, batch OCR, scan_for_ocr_needed
    ├── protect.sh     # encrypt, decrypt, text watermark, PDF stamp, page numbers
    └── repair.sh      # repair (3 strategies), validate, optimize
```

### Descripción de módulos

| Archivo             | Responsabilidad única                                         |
| ------------------- | ------------------------------------------------------------- |
| `main.sh`           | Cargar módulos, trapear limpieza, despachar la operación      |
| `config.sh`         | Variables globales, colores ANSI, rutas conocidas del usuario |
| `lib/logger.sh`     | Formateo y escritura de mensajes de log con niveles           |
| `lib/validator.sh`  | Validar PDFs, dependencias, rangos de página, tamaños         |
| `lib/cli.sh`        | Parsear flags, mostrar ayuda, menú interactivo, prompts       |
| `lib/compress.sh`   | Comprimir con Ghostscript / ocrmypdf, lote, estadísticas      |
| `lib/manipulate.sh` | Operaciones estructurales qpdf: merge/split/extract/rotate    |
| `lib/convert.sh`    | Convertir PDFs a imágenes/texto/HTML y viceversa              |
| `lib/metadata.sh`   | Leer y escribir metadatos XMP y DocInfo con exiftool          |
| `lib/ocr.sh`        | Aplicar OCR con ocrmypdf, detectar PDFs sin texto             |
| `lib/protect.sh`    | Cifrado AES-256, marcas de agua, sellos, numeración           |
| `lib/repair.sh`     | Reparar, validar y optimizar estructura de PDFs               |

---

## Bugs corregidos

### Bug #1: `2>/dev/null` silenciaba errores útiles de Ghostscript y ocrmypdf

**Dónde:** `compress_pdf.sh` v2.0, funciones `compress_with_gs()` y `compress_with_ocr()`  
**Problema:** El stderr de ambos comandos se descartaba incondicionalmente, imposibilitando el diagnóstico cuando un PDF estaba corrupto o protegido.  
**Corrección:** Se captura el stderr en una variable y se logea con `log_debug` únicamente cuando `--verbose` está activo.

### Bug #2: Contadores del lote se perdían en subshells

**Dónde:** Bucle `while IFS= read -r -d ''` en el modo batch de `compress_pdf.sh`  
**Problema:** En ciertas versiones de Bash, el bucle corre en un subshell que no propaga variables al proceso padre; `TOTAL_ORIGINAL` y `TOTAL_COMPRESSED` terminaban en 0.  
**Corrección:** Las estadísticas se acumulan en un archivo temporal `/tmp/pdfsuite_stats_$$` y se leen desde él al terminar el bucle, usando `< <(find ...)` (process substitution) que no crea subshell.

### Bug #3: `qpdf --show-object` no existe

**Dónde:** `arch_pdf_metadata_commands.sh`, sección de qpdf  
**Problema:** El comando `qpdf --show-object=1 archivo.pdf` falla con "invalid option"; esa bandera no existe en qpdf.  
**Corrección:** Se reemplaza por `qpdf --check archivo.pdf` (validación estructural) y `qpdf --json=latest archivo.pdf | grep ...` para inspeccionar metadatos.

### Bug #4: `mutool show trailer/Info` no apunta a metadatos XMP

**Dónde:** `arch_pdf_metadata_commands.sh`, sección de mutool  
**Problema:** `trailer/Info` apunta al diccionario DocInfo clásico (autor, título en formato antiguo), NO al paquete XMP incrustado. El comentario en el script original era incorrecto.  
**Corrección:** Se documenta la distinción y se agrega la ruta correcta para XMP: `mutool show archivo.pdf trailer/Root/Metadata`.

### Bug #5: Bucles `find | while read` sin `IFS= -r`

**Dónde:** Bucles de `arch_pdf_metadata_commands.sh` secciones 7 y 8  
**Problema:** `while read f` sin `IFS=` ni `-r` "come" backslashes y recorta espacios al inicio/fin de nombres de archivo, causando errores con rutas como `~/Zotero/storage/ABC/Mi Artículo.pdf`.  
**Corrección:** Todos los bucles usan `while IFS= read -r f` consistentemente.

---

## Solución de problemas

### "command not found: pdf-suite"

```bash
# Verificar que el wrapper existe
ls -la /usr/local/bin/pdf-suite

# Si no existe, crearlo manualmente
sudo tee /usr/local/bin/pdf-suite > /dev/null << 'EOF'
#!/usr/bin/env bash
exec "${HOME}/Documents/scripts_for_linux/pdf-suite/main.sh" "$@"
EOF
sudo chmod +x /usr/local/bin/pdf-suite
```

### "Permission denied" al ejecutar

```bash
chmod +x ~/Documents/scripts_for_linux/pdf-suite/main.sh
chmod +x ~/Documents/scripts_for_linux/pdf-suite/lib/*.sh
```

### Error "not authorized" en ImageMagick al convertir PDF

```bash
sudo nano /etc/ImageMagick-6/policy.xml
# Cambiar: <policy domain="coder" rights="none" pattern="PDF" />
# Por:     <policy domain="coder" rights="read|write" pattern="PDF" />
```

### PDF protegido con contraseña desconocida

```bash
# Intento con contraseña vacía (PDFs con cifrado débil)
pdf-suite decrypt --password "" documento_protegido.pdf
```

### OCR falla con "no languages found"

```bash
# Instalar idioma español para Tesseract
sudo apt install tesseract-ocr-spa

# Verificar idiomas disponibles
tesseract --list-langs
```

### El PDF reparado pierde marcadores o formularios

Esto es esperado cuando se usa la Estrategia 3 de reparación (re-renderizado con Ghostscript). Es el último recurso: convierte el PDF a imágenes y las re-ensambla, perdiendo la capa de texto e interactividad. Si el PDF tiene marcadores importantes, prueba primero solo con qpdf:

```bash
qpdf --linearize documento_roto.pdf documento_reparado.pdf
```

---

## Cómo contribuir

### Agregar una nueva operación

1. Crea `lib/mi_operacion.sh` con función `run_mi_operacion()` siguiendo el patrón de los módulos existentes.
2. En `main.sh`, agrégalo al bloque `source`:
   ```bash
   source "${_MAIN_DIR}/lib/mi_operacion.sh"
   ```
3. En el bloque `case "$operation"` de `main()`:
   ```bash
   mi-operacion) run_mi_operacion "${op_args[@]}" ;;
   ```
4. En `lib/cli.sh`, agrega la opción al menú en `show_interactive_menu()`.
5. Documenta la nueva operación en este README.

### Estándares de código

- Máximo 30 líneas por función
- Nombres en inglés técnico: `verb_noun()` (ej: `compress_single_pdf`)
- Comenta el **por qué**, no el qué
- Usa `validate_pdf_file` y `validate_output_path` antes de cualquier operación
- Toda salida visible pasa por `log_success / log_failure / log_skip`
- Prueba casos de error: archivo inexistente, sin permisos, PDF corrupto

---

## Notas y advertencias

### PDF_SEARCH_PATHS en config.sh

El archivo `config.sh` incluye todas las carpetas del usuario en `~/Documents/` como rutas de búsqueda sugeridas en el menú interactivo. Estas rutas no afectan el funcionamiento CLI; son solo sugerencias visuales.

### cpdf: licencia AGPL

`cpdf` (Coherent PDF) es gratuito para uso no comercial. Para uso académico y personal en UNSCH es perfectamente válido. Para uso comercial o redistribución consultar la licencia en su repositorio.

### Modo dry-run y archivos temporales

Con `--dry-run` no se escribe ningún archivo de salida, pero sí se crean archivos temporales en `/tmp/pdfsuite_*` que se limpian automáticamente al terminar vía `trap cleanup EXIT`.

### Comportamiento de `--suffix`

El sufijo `_out` por defecto puede cambiarse con `-s`. El script usa el sufijo también para detectar archivos ya procesados y evitar procesarlos de nuevo (prevención de bucles en lote). Si cambias el sufijo, asegúrate de que sea consistente entre llamadas del mismo proyecto.

### `delete` usa Python3

La operación `delete` requiere `python3` para calcular el conjunto inverso de páginas (las que se mantienen). Python3 está disponible por defecto en todas las distribuciones modernas.

---

## Changelog

### v3.0.0 (2026-06) — Refactorización completa + nuevas funcionalidades

- **NUEVO**: Arquitectura modular con 10 archivos lib/
- **NUEVO**: Operaciones merge, split, extract, rotate, reorder, delete, convert, protect, repair, validate, optimize
- **NUEVO**: Menú interactivo TUI
- **NUEVO**: Modo dry-run global
- **NUEVO**: Logging centralizado con niveles INFO/WARN/ERROR/DEBUG
- **NUEVO**: Wrapper global multiplataforma (apt + pacman)
- **FIX #1**: stderr de Ghostscript/ocrmypdf capturado, no descartado
- **FIX #2**: Estadísticas del lote preservadas fuera de subshells
- **FIX #3**: `qpdf --show-object` reemplazado por opciones existentes
- **FIX #4**: Distinción correcta DocInfo vs XMP en mutool
- **FIX #5**: Bucles find usan `IFS= read -r` para nombres con espacios

### v2.0.0 (2026-01) — Primera versión funcional (compress + metadata)

- FIX CRÍTICO: compresión ahora reduce el tamaño (v1.0 lo aumentaba)
- Modo recursivo, estadísticas, umbral configurable

### v1.0.0 (2026-01) — Primera versión (con bugs)

- Compresión aumentaba el tamaño en lugar de reducirlo

---

**Autor:** Edison Achalma · [@achalmed](https://github.com/achalmed)  
**Universidad:** Nacional de San Cristóbal de Huamanga, Ayacucho, Perú  
**Licencia:** MIT — compartir y adaptar con atribución  
**Si te fue útil:** ⭐ en GitHub o [invítame un café](https://www.patreon.com/achalmaedison)
