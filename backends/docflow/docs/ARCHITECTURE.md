---
tipo: doc
titulo: Arquitectura de docflow
estado: activo
---
# Arquitectura de docflow

Este documento explica las decisiones de diseño de docflow v3 y cómo extender
el sistema. Léelo antes de modificar el núcleo.

## Principios

1. **Responsabilidad única por módulo.** Ningún archivo mezcla parsing de CLI
   con lógica de conversión, ni logging con estado de sesión.
2. **El núcleo no conoce formatos.** Los formatos se declaran en
   `lib/formats/` y los motores los consultan a través del registry.
   Añadir un formato jamás toca `lib/core/`.
3. **Errores por archivo, nunca por lote.** No se usa `set -e`: el fallo de
   un documento se registra y el lote continúa. El código de salida final
   refleja si hubo fallos (6) o no (0).
4. **Nada destructivo sin verificación.** Un original solo se respalda o
   elimina si su conversión fue exitosa **y** pasó la verificación del
   Validation Engine.
5. **Bash para orquestar, Python (stdlib) para parsear.** Bash coordina
   procesos y archivos; todo lo que requiere parsing estructurado (XML OOXML,
   TOML, Markdown, JSON) vive en `lib/helpers/*.py` sin dependencias externas.

## Capas

```
bin/docflow
  └─ lib/core/bootstrap.sh        carga módulos, resuelve el comando
       ├─ lib/core/*              infraestructura transversal
       ├─ lib/engines/*           motores de conversión
       ├─ lib/formats/*           declaraciones de formatos (registry)
       └─ lib/commands/cmd_*.sh   un archivo por subcomando
```

### El registry (`lib/core/registry.sh`)

Cada formato declara una línea:

```bash
registry_add docx "document" "Word (OOXML)" pandoc soffice odt ""
#            ext   categoría  descripción    md     pdf     odf ms
```

Los campos `md`/`pdf` son **estrategias** que el motor correspondiente sabe
ejecutar; `odf`/`ms` son extensiones destino para los motores bidireccionales.
El descubrimiento de archivos (`finder`), la tabla `docflow formats`, y los
cuatro comandos de conversión derivan todo de aquí.

### El contrato de los motores

Cada tarea de conversión (`md`, `pdf`, `odf`, `office`) expone dos funciones:

```bash
engine_<tarea>_output_ext <ext_entrada>   # → extensión salida ('' = n/a)
engine_<tarea>_convert <entrada> <salida> # → 0 éxito / ≠0 error
```

`dispatch.sh` es el único llamador y aporta todo lo transversal:
descubrimiento, rutas de salida, overwrite/resume, caché, detección de
protegidos, dry-run, hooks, cronometraje, verificación, registro en sesión
y gestión de originales. Un motor nuevo (p. ej. `epub`) solo implementa el
contrato y gana todas esas capacidades gratis.

### Estrategias del Markdown Engine

| Estrategia | Formatos | Ruta |
|---|---|---|
| `pandoc` | docx, odt, rtf, epub, html, tex | pandoc directo con `--extract-media` |
| `via_docx` | doc, dot, dotx, fodt, ott | soffice → docx → pandoc |
| `pptx_parser` | pptx, pptm, ppsx, potx | `helpers/pptx_to_md.py` (OOXML directo) |
| `via_pptx` | ppt, pps, pot, odp, fodp, otp | soffice → pptx → parser |
| `xlsx_parser` | xlsx, xlsm, xltx | `helpers/xlsx_to_md.py` (OOXML directo) |
| `via_xlsx` | xls, ods, fods, ots | soffice → xlsx → parser |
| `csv_parser` | csv, tsv | `helpers/xlsx_to_md.py --csv` |
| `pdf_text` | pdf | (ocrmypdf →) pdftotext |

**Por qué parsers propios:** pandoc no tiene lectores para pptx/xlsx. La
alternativa clásica (soffice → html → pandoc) pierde notas del presentador,
el orden de las diapositivas y produce HTML sucio. Los contenedores OOXML son
zip+XML documentado: parsearlos con `xml.etree` es más fiel y sin dependencias.

**Post-proceso común** (`_md_postprocess`): consolidación de imágenes en
`<base>_files/images/figure-NNN.*` con reescritura de enlaces relativos
(`mdtools.py fix-media`), procesamiento de imágenes (Media Engine), limpieza
conservadora (`mdtools.py clean`) y frontmatter YAML (Metadata Engine).

### LibreOffice headless (`lib/core/soffice.sh`)

Dos problemas conocidos de `soffice --convert-to` en automatización:

1. **Perfil compartido:** dos instancias simultáneas corrompen el perfil de
   usuario. Solución: `-env:UserInstallation=file:///tmp/...` único por
   invocación (es lo que permite `--jobs N` con seguridad).
2. **Cuelgues:** documentos rotos pueden bloquear soffice indefinidamente.
   Solución: `timeout` (300 s por defecto) + verificación de que el archivo
   esperado realmente existe (soffice devuelve 0 aunque no convierta, p. ej.
   con documentos protegidos).

### Paralelización (`lib/core/parallel.sh`)

En lugar del frágil `export -f` de funciones a subshells, cada worker
re-invoca `docflow __worker <tarea> <archivo>`: un subcomando interno que
re-carga el entorno completo. La configuración viaja por variables
`DOCFLOW_*` exportadas; los resultados convergen en el `results.tsv` de la
sesión (append atómico: líneas < PIPE_BUF). El proceso padre dibuja la barra
de progreso monitorizando ese archivo.

### Sesiones y estado

- `~/.local/state/docflow/sessions/<ts>/results.tsv` — una línea por archivo:
  `status task input output ms bytes_in bytes_out mensaje`.
- Estados: `ok | skip | cached | fail | protected | dry-run`.
- `report.py` genera HTML/CSV/JSON/MD desde ese TSV; `--json` lo emite crudo.
- Se conservan las últimas 50 sesiones.

### Caché (`lib/core/cache.sh`)

Clave = `sha256(contenido) + tarea + opciones relevantes` (flavor, formato de
imagen, motor PDF). Valor = ruta de salida. Hit solo si la salida registrada
coincide con la esperada y aún existe. El índice es append-only (la última
entrada gana); `docflow cache prune` lo compacta.

## Cómo extender

**Formato nuevo:** un archivo en `lib/formats/` con `registry_add`. Si
necesita una estrategia nueva, añádela al `case` del motor correspondiente
y documenta aquí.

**Comando nuevo:** `lib/commands/cmd_<nombre>.sh` con una función
`cmd_<nombre>()` (y opcionalmente `cli_help_<nombre>()`). El bootstrap lo
descubre por convención de nombres; no hay lista central que mantener.

**Motor nuevo:** implementa el contrato de dos funciones y añade la tarea a
`_dispatch_exts_for_task`. Todo lo transversal es gratis.

**Opción de configuración nueva:** valor por defecto en
`config/defaults.conf`, mapeo TOML en `KEYMAP` de `lib/core/config.sh`,
flag en `lib/core/cli.sh` y plantilla en `config_write_template`.

## Convenciones

- `bash -n` limpio y ShellCheck (nivel warning) en CI.
- Funciones privadas de módulo con prefijo `_`.
- Toda salida al usuario pasa por `logger.sh` y va a **stderr**; stdout queda
  reservado para datos (`--json`, `config path`).
- Rutas siempre entrecomilladas; iteración de archivos con NUL
  (`find -print0` / `read -d ''`).
- Mensajes de cara al usuario en español; identificadores en inglés técnico.
