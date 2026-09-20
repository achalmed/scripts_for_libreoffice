---
tipo: changelog
estado: activo
---
# CHANGELOG — scripts_document_studio (DocFlow Studio; repo `scripts_for_libreoffice`)

Cambios del repo y de sus cuatro suites, con fecha ISO y lo más reciente arriba. Versiones vigentes:
docflow **3.0.0**, pdf-suite **3.0.0**, page-counter **2.0.0** (`SCRIPT_VERSION` en su `config.py`),
DocFlow Studio **1.0.0** (`APP_VERSION` en `studio/__init__.py`). Hasta 2026-09-20 este archivo era el
changelog de docflow y vivía en `backends/docflow/`; la entrada v3.0.0 se conserva íntegra al final.

## 2026-09-20 — DOC6: documentación de la suite bajo NORMATIVA §15

- `README.md` reescrito con un solo H1 (`# scripts_document_studio/ — …`), los bloques generados intactos y
  las secciones `Qué es · Uso · Deduplicación · Estructura · Documentación · Límite honesto`; lo que era la
  GUI pasa a `studio/README.md` y la historia v1–v4, a este changelog.
- Nuevos `CLAUDE.md` (+ `AGENTS.md`), `docs/README.md` (generado), `LICENSE` (MIT, que el README declaraba)
  y este `CHANGELOG.md` retomado desde el `git log`.
- `backends/docflow/docs/ARCHITECTURE.md` → `docs/arquitectura.md`: sus principios valen para las cuatro
  suites; docflow conserva un puntero.
- Los tres README de backend llevan frontmatter, H1 §9.6 y `## Límite honesto`.
- `.github/workflows/ci.yml` y `.gitignore` corregidos: sus rutas empezaban por `docflow-studio/`, una
  carpeta que dejó de existir cuando el repo entero pasó a ser la app (el CI no se disparaba).
- Los bloques `suite:`/`suites:` de los README, regenerados por `core/suites.py generar`.

## 2026-09-15 — M5: identidad y normativa de archivos

- `suite.yml` con `id` y línea de identidad en las cuatro suites; identidad §6.2 en los `.sh`/`.py`;
  frontmatter `tipo`/`estado` en `ARCHITECTURE.md` y `ejemplos.md`.
- `studio/`: identidad del paquete por su ruta (A02).
- docflow: `office_meta.py` emite el frontmatter de `meta/NORMATIVA_ARCHIVOS.md` (`tipo: original`,
  claves `snake_case` en español, `id` y `estado`).

## 2026-09-07 — FS1–FS3: contrato de suites

- FS1: un `suite.yml` por suite (`document_studio`, `docflow`, `page_counter`, `pdf_suite`) y bloques de
  README generados por `core/suites.py`.
- FS2: núcleo compartido en `core/` (`shell-lib`, `py-common`); los loggers de pdf-suite y page-counter
  pasan a ser envoltorios; sin rutas literales. docflow conserva su logger (excepción documentada).
- FS3: suites al patrón `main` + `config` + `lib`; `main.py` de la raíz equivale a `python3 -m studio`.

## 2026-09-06 — page-counter 2.0: los blogs viven en el hub

- Los `pub_*` se resuelven en `04 index/_pubs` (`DIR_HUB`, `SUBDIR_PUBS` en `config.py`): son submódulos
  del hub `website-achalma` desde la reorganización del workspace.

## 2026-07-13 — v4: DocFlow Studio 1.0.0

- Aplicación de escritorio PySide6 que centraliza docflow, pdf-suite y page-counter como backends, sin
  reescribirlos: dominios Conversión, PDF, Metadatos, Reportes, Vigilancia y Sistema; explorador,
  consola en vivo, panel de tareas cancelables, visor de logs, QSettings; `install.sh` con `.venv/` y
  entrada de escritorio; `run.sh`.
- Deduplicación: toda operación PDF pasa por pdf-suite; docflow conserva solo PDF/A.

## 2026-07-10 — v3: docflow 3.0.0 y retirada de script_doc_suite

- docflow v3: núcleo (bootstrap, registry de formatos, dispatch, CLI), motores de conversión y helpers
  Python, comandos, instalador multi-distro, tests Bats, CI y autocompletado; README, arquitectura,
  changelog y licencia. Detalle en la entrada v3.0.0 de abajo.
- `script_doc_suite` (v2) retirada, sustituida por docflow; el workflow de CI pasa a la raíz del repo.

## 2026-06 — pdf-suite 3.0.0

- Refactorización completa: arquitectura modular (10 módulos en `lib/`), operaciones merge, split,
  extract, rotate, reorder, delete, convert, protect, repair, validate, optimize; menú interactivo;
  `--dry-run` global; logging con niveles; instalador apt + pacman; cinco bugs corregidos (stderr de
  Ghostscript/ocrmypdf capturado, contadores fuera de subshells, `qpdf --show-object`, DocInfo vs XMP,
  `IFS= read -r`). Antes: 2.0.0 (2026-01, compresión que sí reduce) y 1.0.0 (2026-01, con bugs).

## 2026-06-22 — v2: script_doc_suite

- Los scripts de conversión se consolidan en una suite modular (`to-odf`, `to-pdf`, `to-md`).

## 2026-06-18 — v1 y primer commit

- Origen: `convert_ms_to_odf.sh` (2024), un script único MS Office → ODF. Las versiones anteriores a
  docflow están en el historial de git.

---

## v3.0.0 — 2026-07-10 (docflow)

Reescritura completa desde cero. docflow pasa de suite de scripts a
herramienta CLI profesional.

### Nuevo

- **Arquitectura por motores** (Markdown, PDF, ODF, Office, Media, Metadata,
  OCR, Validation, Reporting, Watch) sobre un registry de formatos: añadir
  formatos o comandos no toca el núcleo.
- **Parsers OOXML propios** para pptx (notas del presentador, tablas,
  imágenes, orden real de diapositivas) y xlsx (todas las hojas, fechas y
  números bien formateados) — formatos que pandoc no puede leer.
- **Operaciones PDF**: merge, split, rotate, extract, compress, PDF/A,
  encrypt/decrypt (AES-256), watermark, OCR, info.
- **Caché SHA-256**: pasadas incrementales sobre colecciones grandes.
- **Verificación post-conversión** (imágenes existentes, PDFs válidos,
  contenedores íntegros) como requisito para tocar originales.
- **Detección de documentos protegidos** con estado de reporte propio.
- **Reportes** HTML/CSV/JSON/Markdown con estadísticas por formato.
- **Frontmatter YAML** con metadatos reales del documento (autor, fechas,
  palabras clave, empresa, idioma).
- **Media Engine**: conversión de imágenes a png/jpg/webp/avif,
  redimensionado y optimización.
- Configuración TOML (`~/.config/docflow/config.toml`), salida `--json`,
  modo `--quiet`, códigos de salida documentados, hooks pre/post,
  `--resume`, barra de progreso con ETA, autocompletado bash/fish,
  suite Bats (33 tests) y CI multi-distro.

### Corregido respecto a v2

- pptx→md ahora funciona de verdad (v2 intentaba pandoc sobre odp, que no
  tiene lector: fallaba siempre).
- LibreOffice en paralelo ya no corrompe el perfil de usuario (perfil
  aislado por invocación) y no puede colgar el lote (timeout).
- Los fallos de conversión ya no dejan salidas parciales que bloqueen el
  reintento.
- soffice devolviendo 0 sin convertir (documentos protegidos) ya no se
  reporta como éxito.
