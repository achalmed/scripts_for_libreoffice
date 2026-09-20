# docflow

<!-- suite:inicio -->
**Suite `docflow`** · objetivo *documentos* · estado *activo* · bash · interfaz cli

Flujos de conversión y ofimática (LibreOffice, pandoc) para documentos de trabajo; herramienta en bin/ con instalador.

- Escribe en: archivos · simula por defecto: no
- Nota: config en config/defaults.conf y lib/core/config.sh; logger propio con niveles y modo quiet (excepción documentada a core/shell-lib)

Comandos:

```bash
main.sh --help
main.sh convert <archivo> --to pdf
bin/docflow …           # mismo programa
```

<sub>Bloque generado desde `suite.yml` por `core/suites.py generar` (2026-09-20); no se edita a mano.</sub>
<!-- suite:fin -->

> **Suite profesional de conversión documental para Linux.**
> Convierte masivamente documentos Microsoft Office, LibreOffice/OpenDocument,
> PDF, EPUB, HTML y LaTeX hacia **Markdown**, **PDF** y entre **Office ↔ ODF** —
> con verificación automática, caché, paralelización y reportes.

```
docflow to-md tesis/          # toda la carpeta → Markdown (imágenes incluidas)
docflow to-pdf informe.docx   # mejor motor disponible según el documento
docflow to-odf legacy/        # docx→odt, xlsx→ods, pptx→odp (y plantillas)
docflow pdf merge libro.pdf cap1.pdf cap2.pdf
```

---

## Tabla de contenidos

- [Los tres pilares](#los-tres-pilares)
- [Instalación](#instalación)
- [Uso esencial](#uso-esencial)
- [Comandos](#comandos)
- [Formatos soportados](#formatos-soportados)
- [Conversión a Markdown en detalle](#conversión-a-markdown-en-detalle)
- [Operaciones PDF](#operaciones-pdf)
- [Selección de archivos](#selección-de-archivos)
- [Seguridad de los originales](#seguridad-de-los-originales)
- [Caché y reanudación](#caché-y-reanudación)
- [Reportes](#reportes)
- [Modo watch](#modo-watch)
- [Hooks](#hooks)
- [Configuración](#configuración)
- [Códigos de salida](#códigos-de-salida)
- [Arquitectura](#arquitectura)
- [Tests](#tests)
- [Solución de problemas](#solución-de-problemas)

---

## Los tres pilares

1. **Conversión a Markdown** — la prioridad del proyecto. Markdown limpio y
   fiel, pensado para Obsidian, Logseq, MkDocs, Hugo, Quarto y GitHub:
   encabezados, listas, tablas, enlaces, notas al pie, ecuaciones, imágenes
   extraídas y ordenadas, notas del presentador y metadatos como frontmatter YAML.
2. **Conversión a PDF** — el mejor motor disponible según el documento
   (LibreOffice para Office/ODF, pandoc+XeLaTeX para Markdown/LaTeX), más un
   juego completo de operaciones: unir, dividir, comprimir, PDF/A, cifrar,
   marca de agua, OCR.
3. **Office ↔ OpenDocument** — bidireccional, incluidas plantillas:
   `docx↔odt`, `xlsx↔ods`, `pptx↔odp`, `dotx↔ott`, `xltx↔ots`, `potx↔otp`.

Dos ventajas técnicas sobre un simple `pandoc`/`soffice` en un bucle:

- **Pandoc no puede leer pptx ni xlsx.** docflow incluye parsers OOXML propios
  (Python stdlib, sin dependencias) que extraen texto, listas anidadas, tablas,
  imágenes y **notas del presentador** de presentaciones, y **todas las hojas**
  de un libro de cálculo con fechas y números bien formateados.
- **LibreOffice headless corrompe su perfil al correr en paralelo.** docflow
  aísla cada invocación en un perfil temporal propio y con timeout, lo que
  permite paralelizar con seguridad (`--jobs`).

## Instalación

```bash
git clone https://github.com/achalmed/docflow.git
cd docflow
./install.sh
```

El instalador detecta tu distro (Arch, Debian/Ubuntu, Fedora, openSUSE),
ofrece instalar dependencias, enlaza `docflow` en el PATH e instala el
autocompletado. Sin permisos de administrador: el enlace va a `~/.local/bin`.

Verifica el estado de las dependencias en cualquier momento:

```bash
docflow doctor
```

| Dependencia         | Rol                                          |               |
| ------------------- | -------------------------------------------- | ------------- |
| pandoc              | Markdown y documentos de texto               | **requerida** |
| python3 (≥3.11)     | parsers OOXML, limpieza, config TOML         | **requerida** |
| libreoffice         | Office↔ODF, presentaciones, hojas de cálculo | **requerida** |
| qpdf, ghostscript   | operaciones PDF, compresión, PDF/A           | opcional      |
| poppler (pdftotext) | PDF → Markdown                               | opcional      |
| imagemagick         | conversión/optimización de imágenes          | opcional      |
| ocrmypdf, tesseract | OCR de PDFs escaneados                       | opcional      |
| inotify-tools       | modo watch instantáneo                       | opcional      |
| texlive-xetex       | pandoc → PDF de alta calidad                 | opcional      |

## Uso esencial

```bash
# Un archivo
docflow to-md tesis/capitulo1.docx

# Una carpeta completa (recursivo)
docflow to-md ~/Documentos/investigacion/

# Un disco entero, excluyendo carpetas, en paralelo, con reporte
docflow to-md /mnt/Datos --exclude backup --exclude privado -j 8 --report html

# Simular primero (siempre recomendable en lotes grandes)
docflow to-md ~/Documentos/ -n -v
```

Por defecto el resultado queda **junto al original** (`cap1.docx` → `cap1.md` +
`cap1_files/images/`), sin sobreescribir nada que ya exista.

## Comandos

| Comando     | Función                                                                               |
| ----------- | ------------------------------------------------------------------------------------- |
| `to-md`     | Office/ODF/PDF/EPUB/HTML → Markdown                                                   |
| `to-pdf`    | Office/ODF/Markdown/HTML/LaTeX → PDF                                                  |
| `to-odf`    | MS Office → OpenDocument                                                              |
| `to-office` | OpenDocument (y Markdown) → MS Office                                                 |
| `pdf <op>`  | merge, split, rotate, extract, compress, pdfa, encrypt, decrypt, watermark, ocr, info |
| `watch`     | vigilar directorio y convertir al vuelo                                               |
| `report`    | reporte de la última sesión (html/csv/json/md)                                        |
| `index`     | índice global + árbol de documentos Markdown                                          |
| `doctor`    | diagnóstico de dependencias con instrucciones por distro                              |
| `formats`   | tabla de formatos soportados                                                          |
| `cache`     | stats / clear / prune del caché de conversiones                                       |
| `config`    | init / show / path de la configuración                                                |

Ayuda contextual: `docflow <comando> --help`.

## Formatos soportados

`docflow formats` imprime la tabla completa. Resumen:

- **Documentos:** docx, docm, doc, dot, dotx, rtf, odt, fodt, ott
- **Presentaciones:** pptx, pptm, ppsx, potx, ppt, pps, pot, odp, fodp, otp
- **Hojas de cálculo:** xlsx, xlsm, xltx, xls, ods, fods, ots, csv, tsv
- **Otros:** pdf, epub, html, txt, tex, md

Añadir un formato nuevo = crear un archivo de ~3 líneas en `lib/formats/`
(ver [Arquitectura](#arquitectura)); el núcleo no se toca.

## Conversión a Markdown en detalle

```bash
docflow to-md documento.docx
```

produce:

```
documento.docx        ← original intacto
documento.md          ← Markdown con frontmatter YAML
documento_files/
    images/
        figure-001.png
        figure-002.png
```

- **Metadatos**: título, autor, fechas, idioma, palabras clave, empresa…
  extraídos del XML del documento e inyectados como frontmatter
  (desactivable con `--no-metadata`).
- **Imágenes**: extraídas, renombradas secuencialmente y enlazadas con rutas
  relativas. `--img-format webp --img-quality 85` las convierte;
  `--img-max-width 1200` las redimensiona; `--img-optimize` las comprime.
- **Presentaciones**: una sección `#` por diapositiva separada con `---`,
  con listas anidadas, tablas y notas del presentador como cita
  (`--no-notes` para omitirlas).
- **Hojas de cálculo**: una sección `##` por hoja, como tabla Markdown.
- **PDF**: texto extraído con poppler; `--ocr` aplica OCR antes
  (PDFs escaneados) con `--ocr-lang spa+eng`.
- **Limpieza** (`--no-clean` para desactivar): HTML residual, atributos de
  pandoc, líneas en blanco duplicadas, normalización UTF-8 (NFC).
- **Destino** con `--flavor`: `gfm` (defecto), `obsidian`, `logseq`,
  `mkdocs`, `hugo`, `quarto`.
- **Verificación automática**: el `.md` existe, tiene contenido y todas las
  imágenes enlazadas existen; si algo falla, se marca como error y el
  original queda a salvo.

## Operaciones PDF

```bash
docflow pdf merge tesis.pdf cap1.pdf cap2.pdf cap3.pdf
docflow pdf split escaneo.pdf 10                # trozos de 10 páginas
docflow pdf extract libro.pdf --pages 5-20 -o capitulo.pdf
docflow pdf rotate apaisado.pdf --degrees 90 --pages 2-4
docflow pdf compress pesado.pdf -o ligero.pdf   # solo si realmente reduce
docflow pdf pdfa articulo.pdf                   # PDF/A-2b para archivado
docflow pdf encrypt privado.pdf --password clave
docflow pdf decrypt privado.pdf --password clave
docflow pdf watermark borrador.pdf --text "BORRADOR"
docflow pdf ocr escaneado.pdf -o buscable.pdf
docflow pdf info documento.pdf
```

En `to-pdf`, `--pdfa` y `--compress` aplican el post-proceso a cada PDF
generado; `--pdf-engine` fuerza un motor concreto.

## Selección de archivos

```bash
-f docx,odt             # solo estas extensiones
--include "*.final.*"   # solo archivos que casen con el glob
-e node_modules -e .git # excluir directorios por nombre (repetible)
-e "*.tmp"              # excluir por patrón
--exclude-regex 'borrador|v[0-9]+' # excluir rutas por regex
--no-default-excludes   # no aplicar la lista por defecto
```

Exclusiones por defecto: `.git node_modules __pycache__ _backup _site
_freeze .quarto .obsidian` y similares. Los archivos de bloqueo de Office
(`~$doc.docx`, `.~lock.*`) se ignoran siempre.

## Seguridad de los originales

Invariante del proyecto: **un original solo se toca si su conversión fue
exitosa Y verificada.**

```bash
--backup     # mueve originales exitosos a _backup/ (nunca pregunta)
--delete     # elimina originales exitosos (pide confirmación)
--force      # suprime la confirmación (automatización/cron)
```

Los archivos con error, protegidos por contraseña o sin convertir **jamás**
se mueven ni eliminan. Los documentos protegidos se detectan antes de
convertir y se reportan con estado propio (`protected`).

## Caché y reanudación

docflow guarda un índice SHA-256 de cada conversión exitosa
(`~/.cache/docflow/`). Un archivo sin cambios con las mismas opciones no se
reconvierte, lo que hace las pasadas incrementales sobre colecciones grandes
casi instantáneas.

```bash
docflow to-md /mnt/Datos --resume   # reanudar un lote interrumpido
docflow to-md tesis/ --no-cache -w  # forzar reconversión
docflow cache stats|prune|clear
```

## Reportes

```bash
docflow to-md tesis/ --report html,json   # al terminar el lote
docflow report md                         # de la última sesión
```

Incluyen: archivos convertidos, errores con causa, protegidos, tiempo total
y promedio, tamaños de entrada/salida, porcentaje de éxito y desglose por
formato. Para integración con scripts: `--json` emite los resultados de la
sesión por stdout y `-q` silencia el resto.

## Modo watch

```bash
docflow watch ~/Descargas                    # → PDF al detectar archivos
docflow watch --to md -f docx ~/notas/
docflow watch ~/buzon --post-hook 'notify-send docflow "$DOCFLOW_OUTPUT"'
```

Usa inotify si está instalado (detección instantánea, CPU cero); si no,
polling cada `--interval` segundos.

## Hooks

```bash
--pre-hook  'echo "convirtiendo $DOCFLOW_INPUT"'
--post-hook 'git -C $(dirname "$DOCFLOW_OUTPUT") add "$DOCFLOW_OUTPUT"'
```

El comando recibe `DOCFLOW_INPUT`, `DOCFLOW_OUTPUT`, `DOCFLOW_STATUS`
(solo post: `ok`/`fail`) y `DOCFLOW_TASK` como variables de entorno.

## Configuración

```bash
docflow config init    # crea ~/.config/docflow/config.toml comentado
```

Todas las opciones de la CLI tienen su equivalente permanente en TOML
(secciones `[general]`, `[markdown]`, `[images]`, `[pdf]`, `[ocr]`,
`[watch]`, `[reports]`, `[hooks]`, `[originals]`).
Precedencia: **CLI > config.toml > defaults**.

## Códigos de salida

| Código | Significado                             |
| ------ | --------------------------------------- |
| 0      | éxito                                   |
| 1      | error general                           |
| 2      | uso incorrecto de la CLI                |
| 3      | ruta de entrada inválida                |
| 4      | no se encontraron archivos que procesar |
| 5      | dependencia requerida ausente           |
| 6      | al menos una conversión falló           |
| 7      | verificación post-conversión fallida    |
| 8      | configuración inválida                  |
| 130    | interrumpido (Ctrl-C) — usa `--resume`  |

## Arquitectura

```
docflow/
├── bin/docflow             # entrada fina: bootstrap + despacho
├── config/defaults.conf    # valores por defecto (capa 1)
├── lib/
│   ├── core/               # infraestructura transversal
│   │   ├── registry.sh     #   registro de formatos (patrón registry)
│   │   ├── dispatch.sh     #   pipeline común de todo lote
│   │   ├── soffice.sh      #   LibreOffice aislado + timeout
│   │   ├── cache.sh        #   caché SHA-256
│   │   ├── parallel.sh     #   workers vía `docflow __worker`
│   │   └── … (config, logger, finder, session, hooks, backup, progress)
│   ├── engines/            # motores de conversión independientes
│   │   ├── markdown.sh  pdf.sh  odf.sh  office.sh
│   │   ├── pdf_ops.sh   media.sh  metadata.sh
│   │   └── validation.sh  reporting.sh  watch.sh
│   ├── formats/            # un módulo por formato (se auto-registran)
│   ├── commands/           # un módulo por subcomando de la CLI
│   └── helpers/            # Python stdlib: parsers OOXML, limpieza, reportes
├── completions/  tests/  docs/
└── install.sh
```

Cada **motor** implementa un contrato de dos funciones
(`engine_<tarea>_output_ext`, `engine_<tarea>_convert`) y es reutilizable
desde cualquier comando. Cada **formato** declara sus capacidades en una
línea de registro. Detalles y decisiones de diseño en
[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

### Añadir un formato nuevo

```bash
# lib/formats/docbook.sh
registry_add xml "markup" "DocBook XML" pandoc pandoc "" ""
```

Nada más: `finder`, `dispatch`, caché, reportes y verificación lo recogen.

## Tests

```bash
bats tests/          # 33 tests en sandbox aislado
```

CI en GitHub Actions: ShellCheck + suite Bats sobre Ubuntu y Arch Linux.

## Solución de problemas

**«protegido por contraseña o corrupto»** — el documento está cifrado
(o truncado). Para PDF: `docflow pdf decrypt archivo.pdf --password …`.

**LibreOffice tarda o se cuelga con un documento** — hay un timeout de 300 s
por archivo (`DOCFLOW_SOFFICE_TIMEOUT` para ajustarlo); el archivo se marca
como fallo y el lote continúa.

**El PDF no tiene las fuentes correctas** — instala las fuentes del documento
(`fc-list | grep Fuente`); LibreOffice sustituye las que faltan.

**Las imágenes no se convierten a WebP** — requiere ImageMagick
(`docflow doctor` te da el comando exacto para tu distro).

**Quiero ver qué haría sin tocar nada** — `docflow to-md -n -v ruta/`.

---

**Licencia:** MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
