---
tipo: readme
estado: activo
---
# scripts_document_studio/ — DocFlow Studio: PDF y ofimática (repo scripts_for_libreoffice)

<!-- suite:inicio -->
**Suite `document_studio`** · objetivo *documentos* · estado *activo* · python · interfaz gui

Interfaz de escritorio (PySide6) sobre docflow, pdf-suite y page-counter.

- Escribe en: archivos · simula por defecto: no
- Depende de: PySide6
- Nota: config en studio/core/settings.py (QSettings); backends docflow, pdf-suite y page-counter en backends/

Comandos:

```bash
main.py
./run.sh                 # usa el .venv si existe
python3 -m studio
```

<sub>Bloque generado desde `suite.yml` por `core/suites.py generar` (2026-09-20); no se edita a mano.</sub>
<!-- suite:fin -->

<!-- suites:inicio -->
Suites de esta carpeta (4); índice global en `meta/INDICE_SCRIPTS.md`. Patrón: M main · C config · L lib.

| Suite | Carpeta | Objetivo | Escribe en | Simula | Timer | Estado | Patrón |
|---|---|---|---|---|---|---|---|
| `docflow` | [scripts_document_studio/backends/docflow](backends/docflow/) | documentos | archivos | no |  | activo | `MCL` |
| `page_counter` | [scripts_document_studio/backends/page-counter](backends/page-counter/) | documentos | ninguno | sí |  | activo | `MCL` |
| `pdf_suite` | [scripts_document_studio/backends/pdf-suite](backends/pdf-suite/) | documentos | archivos | no |  | activo | `MCL` |
| `document_studio` | [scripts_document_studio](./) | documentos | archivos | no |  | activo | `MCL` |

<sub>Bloque generado desde los `suite.yml` por `core/suites.py generar` (2026-09-20); no se edita a mano.</sub>
<!-- suites:fin -->

## Qué es

Una aplicación de escritorio (DocFlow Studio, PySide6) y los tres motores de línea de comandos que
envuelve, para convertir y gestionar documentos de oficina y PDF en Linux: `docflow` (Bash; conversión
documental masiva → Markdown, → PDF, Office ↔ ODF, con caché, paralelización, vigilancia de carpetas y
reportes), `pdf-suite` (Bash; comprimir, unir, dividir, rotar, OCR, cifrar, marca de agua, reparar,
metadatos) y `page-counter` (Python; cuenta las páginas de los `index.pdf` renderizados por la familia de
blogs Quarto y produce un Excel). Resuelven lo que `pandoc` y `soffice` en un bucle no resuelven: pptx y
xlsx que pandoc no lee, LibreOffice que corrompe su perfil al correr en paralelo, y lotes de miles de
archivos en los que el fallo de uno no debe detener el resto.

Tres nombres para una sola cosa: la carpeta es `scripts_document_studio`; el remoto en GitHub sigue
llamándose `scripts_for_libreoffice` (los nombres de repo no cambiaron al fusionar el workspace con el
vault, 2026-09-06); la suite raíz se llama `document_studio` y la GUI, «DocFlow Studio».

**No es** una biblioteca y no reescribe los backends: cada motor sigue siendo una CLI independiente, con
su `suite.yml`, su README y su instalador (docflow, además, con tests Bats y CI); la GUI los invoca como
procesos o, en el caso de `page-counter`, importa sus módulos tal cual. Depende de `core/` (contrato de
suites) y de PySide6 en un `.venv/` local; no tiene `verdad` propia en `meta/workspace.yml` y no escribe
fuera de la carpeta que se le pasa y de los directorios XDG de cada backend.

## Uso

```bash
./install.sh                                          # una vez: .venv con PySide6, lanzador ~/.local/bin/docflow-studio y entrada de escritorio
docflow-studio                                        # o ./run.sh (usa .venv si existe), python3 main.py, python3 -m studio
backends/docflow/install.sh                           # dependencias del sistema de docflow (pandoc, LibreOffice, qpdf, gs, ocrmypdf…)
backends/pdf-suite/install.sh                         # ídem pdf-suite; la app las diagnostica en Sistema → Diagnóstico
backends/docflow/bin/docflow doctor                   # qué falta y el comando exacto para tu distro
backends/docflow/bin/docflow to-md -n -v ruta/        # simular (-n) una conversión a Markdown antes de aplicarla
backends/docflow/bin/docflow to-md ruta/ --jobs 4 --report html   # lote en paralelo con reporte de sesión
backends/pdf-suite/main.sh                            # menú interactivo; con operación: compress, merge, split, ocr, metadata…
backends/pdf-suite/main.sh -n compress documento.pdf  # -n simula sin escribir
cd backends/page-counter && python3 main.py --help    # conteo de páginas de los blogs (escribe solo su excel_databases/)
bats backends/docflow/tests/                          # 33 tests de docflow en sandbox aislado
pyside6-designer studio/resources/ui/main_window.ui   # editar el shell de la GUI; se carga en tiempo de ejecución
```

Regla de oro: `-n`/`--dry-run` antes de un lote. Solo `page-counter` simula por defecto (no escribe
fuera de su carpeta); `docflow` y `pdf-suite` escriben de verdad salvo que se les pida lo contrario.

## Deduplicación

`docflow pdf` y `pdf-suite` implementaban las mismas operaciones (merge, split, rotate, extract,
compress, encrypt, watermark, ocr, info). En la app **toda la manipulación PDF pasa por pdf-suite** —que
es superconjunto— y docflow queda como motor de *conversión documental*; la única operación PDF que
conserva es **PDF/A**, que pdf-suite no implementa. Ninguna operación aparece en dos módulos. La
duplicación de código sigue existiendo en los backends (ambos son CLIs completas); lo que no se duplica
es la puerta de entrada.

## Estructura

| carpeta | qué es | dueño / generador |
|---|---|---|
| `main.py` · `run.sh` · `install.sh` | entrada de la GUI (equivale a `python3 -m studio`); lanzador que usa `.venv/` si existe; instalador del frontend (venv, ~/.local/bin/docflow-studio, `.desktop`) | a mano |
| `studio/` | la aplicación PySide6: `app.py`, `core/` (rutas, QSettings, tareas), `services/` (GUI → línea de comandos exacta), `ui/` (ventana, páginas por dominio, widgets), `studio/resources/ui/main_window.ui` | a mano; `studio/README.md` |
| `backends/docflow/` | CLI Bash de conversión documental (v3): `bin/docflow`, `lib/{core,engines,formats,commands,helpers}`, `tests/` Bats, `completions/`, `install.sh`, `LICENSE` | a mano; `backends/docflow/README.md` |
| `backends/pdf-suite/` | CLI Bash de manipulación PDF (v3.0): `main.sh` + `config.sh` + `lib/` (10 módulos), `install.sh` | a mano; `backends/pdf-suite/README.md` |
| `backends/page-counter/` | Python: `main.py` + `config.py` + `lib/`; `ejemplos.md`; escribe Excel en su `excel_databases/` (ignorado) | a mano; `backends/page-counter/README.md` |
| `suite.yml` (raíz y uno por backend) | manifiesto de cada suite (`core/suite.schema.yml`) | a mano; los bloques de README los genera `core/suites.py generar --aplicar` |
| `docs/` | documentación permanente del repo: `arquitectura.md` (principios comunes a las cuatro suites y diseño de docflow) | a mano; `docs/README.md` lo genera `core/docs.py indice` |
| `requirements.txt` | PySide6, pypdf, openpyxl (frontend y page-counter); las dependencias del sistema las instalan los backends | a mano |
| `.github/workflows/ci.yml` | ShellCheck + Bats de docflow en Ubuntu y Arch (se dispara solo con cambios en `backends/docflow/`) | a mano |
| `CHANGELOG.md` · `LICENSE` | versiones con fecha ISO de las cuatro suites; MIT | a mano |
| `.venv/` | entorno del frontend creado por `install.sh` (~660 MB; ignorado) | `install.sh` |

## Documentación

| documento | para qué leerlo |
|---|---|
| `CLAUDE.md` | reglas para el asistente: invariantes de diseño, cómo verificar, trampas |
| `docs/arquitectura.md` | los cinco principios (errores por archivo, Bash orquesta y Python parsea…) y el diseño de docflow: registry, motores, soffice aislado, paralelización, sesiones, caché |
| `studio/README.md` | la GUI: dominios, arquitectura por capas, reglas para añadir operaciones y dominios, Qt Designer |
| `backends/docflow/README.md` | manual de docflow: comandos, formatos, seguridad de originales, caché, watch, hooks, códigos de salida |
| `backends/pdf-suite/README.md` | manual de pdf-suite: operaciones, dependencias, notas y advertencias |
| `backends/page-counter/README.md` · `backends/page-counter/ejemplos.md` | manual y casos de uso del contador de páginas |
| `CHANGELOG.md` | qué versión tiene cada herramienta y desde cuándo |
| `meta/INDICE_SCRIPTS.md` | las 4 suites entre las del workspace (generado) |

## Límite honesto

- **Solo docflow tiene pruebas** (33 tests Bats, ShellCheck y CI); `pdf-suite`, `page-counter` y la GUI se
  comprueban a mano: `bash -n`, `py_compile`, `--dry-run` y abrir la app.
- **La GUI no simula**: ejecuta la línea de comandos exacta que muestra en la consola; la simulación es la
  de cada backend (`-n`). Una operación a la vez por tarea; las tareas corren en `QProcess`/`QThread` y se
  pueden cancelar.
- **Los originales solo los toca docflow y solo si la conversión fue exitosa y verificada** (`--backup`,
  `--delete`); los archivos con error o protegidos por contraseña jamás se mueven ni eliminan.
- **PDF/A vive en docflow; todo lo demás del PDF, en pdf-suite**: no hay una sola CLI para todo, y
  `pdf-suite` sigue implementando lo que `docflow pdf` también implementa.
- **docflow tiene logger propio** (niveles y modo `--quiet`), excepción documentada al logger de
  `core/shell-lib` en su `suite.yml`; `page-counter` resuelve la raíz de los blogs con `Path.home()` en su
  `config.py`, no con `core/env.py`.
- **Estado fuera del repo**: config TOML de docflow en `~/.config/docflow/`, sesiones y caché en
  ~/.local/state/docflow/ y ~/.cache/docflow/ (se conservan 50 sesiones), log de la GUI en
  ~/.local/state/docflow-studio/studio.log, ajustes en QSettings (~/.config/achalma/docflow-studio.conf).
- **Lo que no hace**: no edita documentos (convierte, manipula, cuenta), no sincroniza con Calibre ni
  Zotero (eso es `scripts_for_calibre`), no compila LaTeX (`scripts_for_latex`).
- Licencia MIT (`LICENSE`; docflow conserva la suya en `backends/docflow/LICENSE`).
