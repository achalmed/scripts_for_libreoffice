---
tipo: doc
titulo: Referencia de pdf-suite, la CLI de manipulación de PDF
estado: activo
---
# Referencia de pdf-suite

Sintaxis, opciones, operaciones y dependencias de `backends/pdf-suite/main.sh`. Para instalarlo y
empezar, `backends/pdf-suite/README.md`; para su estructura interna y cómo añadir una operación,
`docs/arquitectura.md`. La ayuda de la propia CLI (`main.sh --help`) manda sobre este documento si
discrepan. En los ejemplos, `pdf-suite` es `backends/pdf-suite/main.sh` o el lanzador que instala
`backends/pdf-suite/install.sh`.

## Dependencias

### Sistema operativo

- Ubuntu 22.04+ / Kubuntu 22.04+ (probado)
- Arch Linux / Archcraft (compatible con ajuste del instalador)
- Cualquier Linux con Bash ≥ 5.0

### Dependencias obligatorias (sin ellas no arranca; `backends/pdf-suite/lib/validator.sh`, `DEPS_REQUIRED`)

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
| `cpdf`       | manual, ver abajo | Marca de agua texto, numeración   |

### cpdf (marca de agua y numeración)

No lo instala `install.sh`: es un binario de Coherent Graphics, gratuito para uso no comercial
(licencia AGPL para el resto), que se descarga de su repositorio `cpdf-binaries` y se pone en el `PATH`.

## Sintaxis y opciones globales

```text
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
| `--log-file`      | Guardar log en ~/.local/share/pdf-suite/logs/          |
| `--version`       | Mostrar versión                                          |
| `-h, --help`      | Mostrar ayuda completa                                   |

## Operaciones detalladas

### `compress` — Comprimir PDFs

```bash
# Un archivo con método ebook (recomendado para lectura)
pdf-suite compress -m ebook libro.pdf

# Toda la carpeta recursivamente
pdf-suite compress -m ebook -r carpeta/

# Solo comprimir si reduce más del 20%
pdf-suite compress -m ebook -r -t 20 carpeta/

# Máxima compresión (para email/web)
pdf-suite compress -m screen informe.pdf
```

| Método     | DPI | Reducción típica | Uso                               |
| ---------- | --- | ---------------- | --------------------------------- |
| `screen`   | 72  | 80–95%           | Web, email, máxima compresión     |
| `ebook`    | 150 | 60–85%        | **Recomendado** — lectura digital |
| `printer`  | 300 | 40–70%           | Para imprimir                     |
| `prepress` | 300 | 20–50%           | Impresión profesional             |
| `ocr`      | —   | 50–80%        | PDFs escaneados                   |

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
pdf-suite ocr -r -l spa escaneos/

# Ver qué PDFs necesitan OCR antes de procesarlos
pdf-suite ocr --scan carpeta/
```

### `metadata` — Ver y editar metadatos

```bash
# Ver información completa de un PDF
pdf-suite info documento.pdf
pdf-suite metadata documento.pdf

# Ver metadatos en JSON
pdf-suite metadata --json documento.pdf

# Escribir metadatos
pdf-suite metadata --set-title "Título del documento" \
                   --set-author "Autor" \
                   --set-subject "Tema" \
                   --set-keywords "economía, Ayacucho, estadística" \
                   documento.pdf

# Título desde nombre de archivo (lote)
pdf-suite metadata --from-filename -r carpeta/

# Ver qué PDFs de una carpeta no tienen metadatos
pdf-suite metadata --check-missing carpeta/
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
  ebook        67.3 MB            56%      Lectura digital
  printer      89.4 MB            42%      Imprimir
  prepress    112.0 MB            27%      Impresión pro
  ocr          58.7 MB            62%      Escaneados
```

### `deps` — Verificar dependencias

```bash
pdf-suite deps
```

## Problemas frecuentes

### «command not found: pdf-suite»

`backends/pdf-suite/install.sh` crea `/usr/local/bin/pdf-suite`, pero hoy ese lanzador apunta a una copia
en una ruta fija ajena al repo (pendiente en `docs/decisiones.md`). Mientras tanto, llama a
`backends/pdf-suite/main.sh` desde el repo, o a un envoltorio propio que haga `exec` sobre él: un enlace
simbólico no sirve, porque `main.sh` busca `lib/` junto a la ruta por la que se le invoca.

### «Permission denied» al ejecutar

```bash
chmod +x backends/pdf-suite/main.sh backends/pdf-suite/lib/*.sh   # desde la raíz del repo
```

### Error «not authorized» en ImageMagick al convertir PDF

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

### OCR falla con «no languages found»

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
