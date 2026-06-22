# docflow v2.0.0

> Suite unificada de conversión de documentos Office para Arch Linux.  
> Convierte `.docx`, `.odt`, `.pptx`, `.odp`, `.xlsx`, `.ods` y más a ODF, PDF o Markdown desde un único comando.

---

## 📋 Tabla de Contenidos

- [Descripción](#descripción)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Uso](#uso)
- [Comandos](#comandos)
- [Arquitectura](#arquitectura)
- [Bugs Corregidos](#bugs-corregidos)
- [Solución de Problemas](#solución-de-problemas)
- [Cómo Contribuir](#cómo-contribuir)
- [Notas y Advertencias](#notas-y-advertencias)

---

## 📖 Descripción

**docflow** fusiona y extiende tres scripts independientes (`doc2md.sh`, `convert_doc_to_odp.sh`, `convert_odt_and_odp_to_pdf_parallel.sh`) en una suite modular con un único punto de entrada.

### Lo que puede hacer

| Comando  | Convierte                            | Notas                                     |
| -------- | ------------------------------------ | ----------------------------------------- |
| `to-odf` | `.docx/.xlsx/.pptx/...` → ODF        | Usa LibreOffice, conserva formato         |
| `to-pdf` | Office/ODF → `.pdf`                  | Pandoc para texto, LibreOffice para resto |
| `to-md`  | `.docx/.odt/.pptx/.odp` → `.md`      | Extrae imágenes ordenadas al lado del .md |
| `watch`  | Vigilar carpeta y convertir al vuelo | inotify (eficiente) o polling (fallback)  |
| `report` | Reporte HTML o CSV de la sesión      | Estadísticas + tabla de conversiones      |

### Flujo típico de trabajo

```
Archivos .docx/.pptx en ~/Documents/trabajos/
        ↓
  docflow to-pdf -i ~/Documents/trabajos/
        ↓
  PDFs en el mismo directorio
        ↓
  docflow report --report-format html
        ↓
  ~/scripts_for_libreoffice/logs/reporte_20240601_120000.html
```

---

## ⚙️ Requisitos

### Sistema Operativo

- Arch Linux / Archcraft / Manjaro (testado con zsh y Fish)
- Compatible con cualquier distro Linux con los paquetes equivalentes

### Dependencias

| Herramienta     | Rol                                    | Requerida      |
| --------------- | -------------------------------------- | -------------- |
| `pandoc`        | Conversión a Markdown y PDF (texto)    | ✅ Sí          |
| `python3`       | Corrección de rutas de imágenes en .md | ✅ Sí          |
| `libreoffice`   | Conversión ODF y PDF (presentaciones)  | ⚠️ Recomendada |
| `imagemagick`   | Optimización de imágenes extraídas     | ⭕ Opcional    |
| `inotify-tools` | Modo watch eficiente basado en eventos | ⭕ Opcional    |

---

## 🚀 Instalación

### Paso 1: Ubicar el proyecto (no se mueve de aquí)

```bash
# El proyecto siempre vive en:
~/Documents/scripts_for_libreoffice/
```

### Paso 2: Ejecutar el instalador

```bash
cd ~/Documents/scripts_for_libreoffice/
chmod +x install.sh
./install.sh
```

El instalador:

1. Instala `pandoc` y `python` vía `pacman`
2. Pregunta si instalar `libreoffice-still`, `imagemagick` e `inotify-tools`
3. Da permisos de ejecución a todos los scripts
4. Crea el symlink `/usr/local/bin/docflow` → `main.sh`

### Paso 3 (alternativo): Sin permisos de sudo

Agrega un alias en `~/.zshrc` o `~/.config/fish/config.fish`:

```bash
# zsh / bash
alias docflow="$HOME/Documents/scripts_for_libreoffice/main.sh"
```

```fish
# Fish
alias docflow "$HOME/Documents/scripts_for_libreoffice/main.sh"
```

---

## 💻 Uso

### Sintaxis general

```
docflow <COMANDO> [OPCIONES] -i <RUTA>
```

La `<RUTA>` puede ser:

- **Un archivo**: `docflow to-pdf -i ~/Descargas/informe.docx`
- **Un directorio**: `docflow to-pdf -i ~/Documents/trabajos/` (procesa recursivamente)

### Opciones globales

| Flag                    | Descripción                                           |
| ----------------------- | ----------------------------------------------------- |
| `-i, --input <ruta>`    | Archivo o directorio a procesar                       |
| `-o, --output <dir>`    | Directorio de salida (por defecto: junto al original) |
| `-f, --formats <lista>` | Formatos separados por coma: `docx,pptx`              |
| `-v, --verbose`         | Mostrar información detallada                         |
| `-n, --dry-run`         | Simular sin escribir ni eliminar nada                 |
| `-w, --overwrite`       | Sobreescribir archivos de destino existentes          |
| `-l, --log <archivo>`   | Guardar log en archivo además de stdout               |
| `--delete`              | Eliminar originales (pide confirmación)               |
| `--backup`              | Mover originales a `_backup/` en vez de eliminar      |
| `--version`             | Mostrar versión                                       |
| `-h, --help`            | Mostrar ayuda completa                                |

---

## 📚 Comandos

### `to-odf` — Convertir a OpenDocument

Convierte archivos MS Office a sus equivalentes ODF:

```bash
# Toda una carpeta
docflow to-odf -i ~/Documents/trabajos/

# Solo .docx y .xlsx
docflow to-odf -i ~/Documents/ -f docx,xlsx

# Con guardado en otra carpeta
docflow to-odf -i ~/Documents/informe.docx -o ~/Documents/odf/

# Sobreescribir existentes y mover originales a _backup/
docflow to-odf -i ~/Documents/ -w --backup

# Simular primero (ver qué convertiría)
docflow to-odf -n -v -i ~/Documents/
```

**Mapa de conversiones:**

| Entrada                                       | Salida |
| --------------------------------------------- | ------ |
| `.docx` / `.doc`                              | `.odt` |
| `.dotx`                                       | `.ott` |
| `.xlsx` / `.xls` / `.xlsm`                    | `.ods` |
| `.xltx`                                       | `.ots` |
| `.pptx` / `.ppt` / `.pptm` / `.ppsx` / `.pps` | `.odp` |
| `.potx`                                       | `.otp` |

---

### `to-pdf` — Convertir directamente a PDF

```bash
# Un archivo específico con nombre y ruta completa
docflow to-pdf -i /home/achalmaedison/Documents/tesis/capitulo1.docx

# Toda una carpeta, PDFs en otra ubicación
docflow to-pdf -i ~/Documents/presentaciones/ -o ~/Escritorio/PDFs/

# Solo presentaciones
docflow to-pdf -i ~/Documents/ -f pptx,odp

# Paralelo automático para >3 archivos, con log
docflow to-pdf -i ~/Documents/tesis/ -l ~/logs/conversion.log

# Eliminar originales tras conversión (pide confirmación)
docflow to-pdf -i ~/Documents/borradores/ --delete
```

---

### `to-md` — Convertir a Markdown

```bash
# Convertir toda la carpeta ideas
docflow to-md -i ~/Documents/ideas/

# Solo .odt y .docx, sobreescribir existentes
docflow to-md -f odt,docx -w -i ~/Documents/ideas/

# Excluir carpetas específicas
docflow to-md -e borradores -e privado -i ~/Documents/ideas/

# Imágenes en JPEG de alta calidad
docflow to-md --img-format jpg --img-quality 95 -i ~/Documents/

# Con pandoc: tabla de contenidos
docflow to-md --pandoc-args '--toc --toc-depth=3' -i ~/Documents/ideas/

# Dry-run para ver qué convertiría
docflow to-md -n -v -i ~/Documents/ideas/
```

**Estructura de salida para `index.odt` con 2 imágenes:**

```
antes:
  ideas/2017-analisis/index.odt

después:
  ideas/2017-analisis/index.odt        ← original intacto
  ideas/2017-analisis/index.md         ← Markdown generado
  ideas/2017-analisis/index_files/
      figure-md/
          fig-001.png
          fig-002.png
```

---

### `watch` — Vigilar y convertir automáticamente

```bash
# Vigilar ~/Descargas y convertir a PDF al detectar nuevos archivos
docflow watch -i ~/Descargas/

# Solo .docx y .pptx, modo verbose
docflow watch -i ~/Descargas/ -f docx,pptx -v

# Con comando post-conversión (notificación de escritorio)
docflow watch -i ~/Descargas/ --watch-cmd "notify-send 'docflow' 'Conversión completada'"

# Polling cada 60 segundos si no hay inotify
docflow watch -i ~/Descargas/ --interval 60
```

**Modos de vigilancia:**

- **inotify** (por defecto si `inotify-tools` está instalado): detección instantánea basada en eventos del kernel, sin uso de CPU en espera.
- **polling** (fallback): escanea el directorio cada `--interval` segundos comparando snapshots.

---

### `report` — Generar reporte de sesión

```bash
# Reporte HTML de la sesión actual
docflow to-pdf -i ~/Documents/ && docflow report --report-format html

# Reporte CSV en ubicación específica
docflow report --report-format csv --report-out ~/logs/sesion.csv

# El reporte automático se genera si usas --report-format al convertir
docflow to-odf -i ~/Documents/ --report-format html
```

---

## 🗂️ Arquitectura

```
~/Documents/scripts_for_libreoffice/
├── main.sh                   # Punto de entrada único — orquesta módulos
├── install.sh                # Instalador interactivo
├── README.md                 # Esta documentación
├── logs/                     # Logs y reportes generados (creado automáticamente)
└── lib/
    ├── config.sh             # Constantes, colores, mapas de conversión
    ├── logger.sh             # Sistema de logging (INFO/WARN/ERROR/DEBUG)
    ├── validator.sh          # Validación de dependencias, rutas y entradas
    ├── cli.sh                # Parsing de argumentos y función show_help()
    ├── converter_odf.sh      # Motor to-odf: MS Office → OpenDocument
    ├── converter_pdf.sh      # Motor to-pdf: Office/ODF → PDF (con paralelo)
    ├── converter_md.sh       # Motor to-md: Office/ODF → Markdown
    ├── watcher.sh            # Modo watch con inotify/polling
    └── reporter.sh           # Generador de reportes HTML y CSV
```

### Descripción de módulos

| Archivo                | Responsabilidad única                                       |
| ---------------------- | ----------------------------------------------------------- |
| `main.sh`              | Bootstrap, carga módulos, despacha comando                  |
| `lib/config.sh`        | Toda la configuración en un solo lugar; sin "magic strings" |
| `lib/logger.sh`        | Formato consistente, niveles, escritura a archivo           |
| `lib/validator.sh`     | Falla rápido con mensajes claros antes de tocar archivos    |
| `lib/cli.sh`           | Toda la CLI en un lugar; agregar flags = cambio aquí solo   |
| `lib/converter_odf.sh` | Convierte MS Office → ODF usando LibreOffice                |
| `lib/converter_pdf.sh` | Convierte a PDF; elige pandoc o soffice según extensión     |
| `lib/converter_md.sh`  | Convierte a Markdown; corrige rutas de imágenes con Python  |
| `lib/watcher.sh`       | Detecta archivos nuevos con inotify o polling               |
| `lib/reporter.sh`      | Genera HTML/CSV con los resultados acumulados de la sesión  |

---

## 🐛 Bugs Corregidos

### Bug #1: `set -euo pipefail` ausente en scripts originales

- **Descripción**: Sin `pipefail`, errores en pipelines pasaban silenciosamente como éxito.
- **Impacto**: Archivos corrompidos o parcialmente convertidos marcados como OK.
- **Corrección**: `main.sh` tiene `set -euo pipefail`; los módulos usan `|| return N` para control de flujo deliberado.

### Bug #2: Array asociativo `CONVERTED_FILES` invisible en subshells

- **Descripción**: `bash` no exporta arrays asociativos a subshells; `delete_originals()` nunca veía las claves.
- **Impacto**: La función de eliminación nunca eliminaba nada (silenciosamente).
- **Corrección**: Reemplazado por archivo temporal de registro compartido entre subshells.

### Bug #3: `parallel` bloquea interactivamente pidiendo cita académica

- **Descripción**: En la primera ejecución, GNU parallel muestra un aviso y espera input.
- **Impacto**: El script se colgaba en entornos no interactivos (cron, CI).
- **Corrección**: Reemplazado por `xargs -P` que es portable y sin bloqueos.

### Bug #4: `LD_LIBRARY_PATH` hardcodeado para Ubuntu, roto en Arch Linux

- **Descripción**: La ruta `/usr/lib/libreoffice/program` no existe en Arch.
- **Impacto**: LibreOffice fallaba silenciosamente en Arch/Archcraft.
- **Corrección**: Se detecta dinámicamente con `find`; en Arch no es necesario (los paquetes ya lo configuran).

### Bug #5: `SCRIPT_DIR` no resolvía symlinks

- **Descripción**: `dirname "${BASH_SOURCE[0]}"` devuelve el directorio del symlink, no del script real.
- **Impacto**: `install.sh` no encontraba `doc2md.sh` cuando se llamaba desde un symlink.
- **Corrección**: Cambiado a `readlink -f "${BASH_SOURCE[0]}"` en todos los scripts.

### Bug #6: `parallel --jobs 4` hardcodeado

- **Descripción**: 4 workers fijos ignoraba sistemas con 2 núcleos (thrashing) o 16 (desperdicio).
- **Impacto**: Degradación de rendimiento en hardware no estándar.
- **Corrección**: `PARALLEL_JOBS="$(nproc)"` en `config.sh`.

---

## 🔧 Solución de Problemas

**Las presentaciones no se convierten correctamente**

```bash
sudo pacman -S libreoffice-still
```

**El modo watch no detecta archivos al instante**

```bash
# Instalar inotify-tools para vigilancia basada en eventos del kernel
sudo pacman -S inotify-tools
```

**`Permission denied` al ejecutar**

```bash
chmod +x ~/Documents/scripts_for_libreoffice/main.sh
chmod +x ~/Documents/scripts_for_libreoffice/lib/*.sh
```

**Las imágenes del Markdown no se convierten a JPEG**

```bash
sudo pacman -S imagemagick
```

**Quiero ver exactamente qué haría sin ejecutar nada**

```bash
docflow to-pdf -n -v -i ~/Documents/trabajos/
```

**El PDF generado no tiene las fuentes correctas**

```bash
# LibreOffice necesita las fuentes instaladas en el sistema
fc-list | grep "NombreFuente"
sudo pacman -S ttf-nombre-fuente
```

---

## 🤝 Cómo Contribuir

### Para agregar un nuevo formato de entrada a `to-odf`:

1. Edita `lib/config.sh` → agrega una línea a `ODF_CONVERSION_MAP`
2. Formato: `"extension_entrada:extension_salida:Descripción"`
3. No hay que tocar ningún otro archivo.

### Para agregar un nuevo comando (ej. `to-epub`):

1. Crea `lib/converter_epub.sh` con una función `run_to_epub()`
2. Agrega el source en `main.sh`
3. Agrega el case en el `main()` de `main.sh`
4. Agrega el flag en `lib/cli.sh` → `validate_command()` y `show_help()`

### Estándares del proyecto

- Máximo 30 líneas por función
- Nombres en inglés técnico: `convert_single_file()`, no `hacer_cosa()`
- Comenta el "por qué", no el "qué"
- Toda salida al usuario pasa por `lib/logger.sh`
- Toda validación va en `lib/validator.sh`

---

## ⚠️ Notas y Advertencias

**Conversión en paralelo**: Se activa automáticamente para lotes de más de 3 archivos en `to-pdf`. Para deshabilitar y procesar secuencialmente, no hay flag específico en esta versión — edita `PARALLEL_JOBS=1` en `lib/config.sh`.

**Modo watch y archivos grandes**: inotifywait dispara `CLOSE_WRITE`, que ocurre cuando la aplicación cierra el archivo. Para copias de red o sincronizadores como Nextcloud, el evento puede llegar antes de que el archivo esté completamente sincronizado. La pausa de 0.5s en `lib/watcher.sh` mitiga esto pero no lo elimina completamente.

**Eliminación de originales**: siempre se pide confirmación interactiva (`--delete`). En entornos no interactivos (cron), usa `--backup` en su lugar, que nunca pregunta.

**Archivos con imágenes vinculadas** (no incrustadas): pandoc no puede extraerlas. El `.md` se genera correctamente pero sin esas imágenes. Ejecuta con `-v` para identificar qué archivos tienen este problema.

**Licencia**: MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
