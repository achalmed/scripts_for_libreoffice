# scripts_for_libreoffice

<!-- suites:inicio -->
Suites de esta carpeta (4); índice global en `meta/INDICE_SCRIPTS.md`. Patrón: M main · C config · L lib.

| Suite | Carpeta | Objetivo | Escribe en | Simula | Timer | Estado | Patrón |
|---|---|---|---|---|---|---|---|
| `docflow` | [scripts_document_studio/backends/docflow](backends/docflow/) | documentos | archivos | no |  | activo | `··L` |
| `page_counter` | [scripts_document_studio/backends/page-counter](backends/page-counter/) | documentos | ninguno | sí |  | activo | `MCL` |
| `pdf_suite` | [scripts_document_studio/backends/pdf-suite](backends/pdf-suite/) | documentos | archivos | no |  | activo | `MCL` |
| `document_studio` | [scripts_document_studio/studio](studio/) | documentos | archivos | no |  | activo | `···` |

<sub>Bloque generado desde los `suite.yml` por `core/suites.py generar` (2026-09-07); no se edita a mano.</sub>
<!-- suites:fin -->

Herramientas para la conversión y gestión de documentos Office/LibreOffice
y PDF en Linux, unificadas en una aplicación de escritorio.



# DocFlow Studio

> **Aplicación de escritorio unificada** (PySide6 + Qt Designer) que centraliza
> todas las herramientas del proyecto `scripts_for_libreoffice`: conversión
> documental, gestión de PDF, metadatos, reportes y vigilancia de carpetas.
> 
> Los scripts existentes **no se reescribieron**: viven en `backends/` y la
> interfaz gráfica los invoca como motores.

## Ejecutar

```bash
./install.sh        # una vez: venv + lanzador + entrada de escritorio
docflow-studio      # o ./run.sh
```

Requisitos del frontend: Python ≥ 3.11 y PySide6 (los instala `install.sh` en
un venv local). Las dependencias del sistema de cada backend (pandoc,
LibreOffice, qpdf, Ghostscript, ocrmypdf…) se instalan con
`backends/docflow/install.sh` y `backends/pdf-suite/install.sh`, y se
diagnostican desde la propia app (**Sistema → Diagnóstico**).

## Dominios funcionales

La app está organizada por **dominios**, no por scripts. Cada función existe
en un solo lugar:

| Dominio                      | Qué ofrece                                                                                                                                                                                                                                                                  | Backend                    |
| ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------- |
| **Conversión de documentos** | Office/ODF/PDF/EPUB/HTML/LaTeX → Markdown · PDF · Office↔ODF, con lotes, paralelización, caché, dry-run y backup de originales                                                                                                                                              | docflow                    |
| **Gestión de PDF**           | Organizar (unir, dividir, extraer, rotar, reordenar, eliminar), optimizar (comprimir, comparar métodos, linearizar, reparar, validar, PDF/A), seguridad (cifrar, descifrar, marca de agua, numerar), contenido (PDF↔imagen/texto/HTML, extraer imágenes, OCR) e información | pdf-suite (PDF/A: docflow) |
| **Metadatos**                | Ver/editar XMP y DocInfo, título desde nombre de archivo en lote, auditoría de metadatos faltantes                                                                                                                                                                          | pdf-suite                  |
| **Reportes**                 | Reporte de la última sesión de conversión (html/md/csv/json) y conteo de páginas de los blogs Quarto con salida Excel                                                                                                                                                       | docflow + page-counter     |
| **Vigilancia**               | Vigilar una carpeta y convertir al vuelo lo que aparezca                                                                                                                                                                                                                    | docflow watch              |
| **Sistema**                  | Diagnóstico unificado de dependencias, caché de conversiones, config.toml de docflow, rutas de backends                                                                                                                                                                     | docflow + pdf-suite        |

### Deduplicación

`docflow pdf` y `pdf-suite` implementaban las mismas operaciones (merge,
split, rotate, extract, compress, encrypt, watermark, ocr, info). En la app
**toda la manipulación PDF pasa por pdf-suite** — que es superconjunto — y
docflow queda como motor de *conversión documental*; la única operación PDF
que conserva es **PDF/A**, que pdf-suite no implementa. Ninguna operación
aparece en dos módulos.

## Experiencia de usuario

- **Navegación lateral** por dominios con páginas apiladas.
- **Explorador de archivos** acoplable: doble clic o «Usar como entrada»
  envía la ruta a la página activa. También hay *drag & drop* directo.
- **Consola integrada**: salida en vivo (streaming) de cada tarea.
- **Panel de tareas**: estado, duración y cancelación (las tareas corren en
  QProcess/QThread — la interfaz **nunca se bloquea**, ni en lotes de miles
  de archivos).
- **Visor de logs**: histórico persistente en
  `~/.local/state/docflow-studio/studio.log`, con filtros.
- **Configuración persistente** (QSettings): geometría, última página,
  últimos directorios, rutas de backends.

## Arquitectura

```
docflow-studio/
├── run.sh / install.sh        # lanzador y instalador del frontend
├── requirements.txt
├── backends/                   # los motores existentes, INTACTOS
│   ├── docflow/                #   CLI Bash de conversión documental (v3)
│   ├── pdf-suite/              #   CLI Bash de manipulación PDF (v3)
│   └── page-counter/           #   Python: conteo de páginas Quarto (v2)
└── studio/                     # la aplicación PySide6
    ├── app.py  __main__.py     # arranque (logging, QApplication)
    ├── core/                   # infraestructura transversal
    │   ├── paths.py            #   localización de backends (+ overrides)
    │   ├── settings.py         #   QSettings (config persistente)
    │   └── tasks.py            #   ProcessTask/PythonTask/TaskManager
    ├── services/               # adaptadores: GUI → línea de comandos exacta
    │   ├── docflow.py          #   conversión, watch, report, doctor, caché, PDF/A
    │   ├── pdfsuite.py         #   TODAS las operaciones PDF
    │   └── pagecounter.py      #   import directo del backend Python
    ├── ui/
    │   ├── main_window.py      # controlador del shell (.ui) + docks
    │   ├── widgets/            # explorador, consola, tareas, log
    │   └── pages/              # una página por dominio funcional
    │       ├── base.py         #   BasePage, ListaRutas, SelectorRuta
    │       ├── conversion.py  pdf.py  metadata.py
    │       └── reports.py  watch.py  system.py
    └── resources/ui/main_window.ui   # shell editable con Qt Designer
```

Reglas de la arquitectura (pensadas para crecer sin tocarla):

1. **Las páginas no conocen los CLIs**: construyen sus parámetros y llaman a
   `services/`, que devuelve la línea de comandos. Un cambio de flags de un
   backend se corrige en un solo archivo.
2. **Los backends no se modifican**: siguen siendo utilizables desde la
   terminal exactamente igual que antes (sus instaladores, tests y CI
   siguen operativos).
3. **Toda operación es una Task**: cualquier módulo nuevo hereda consola,
   panel de tareas, log y cancelación sin escribir nada.
4. **Añadir una operación PDF** = una entrada en `OPERACIONES`
   (`ui/pages/pdf.py`) + una función de una línea en `services/pdfsuite.py`.
5. **Añadir un dominio nuevo** = una página en `ui/pages/` registrada en la
   tupla `PAGINAS` de `main_window.py`.

## Editar la interfaz con Qt Designer

```bash
pyside6-designer studio/resources/ui/main_window.ui
```

El shell (menús, navegación, stack) es un `.ui` estándar; se carga en
tiempo de ejecución con `QUiLoader`, así que los cambios de Designer se ven
al reiniciar la app sin recompilar nada.

---

**Licencia:** MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)

## DocFlow Studio — la aplicación de escritorio

**[docflow-studio](docflow-studio/)** (PySide6) centraliza todas las
capacidades del proyecto en una sola interfaz organizada por dominios:
conversión de documentos, gestión de PDF, metadatos, reportes, vigilancia
de carpetas y diagnóstico del sistema.

```bash
cd docflow-studio && ./install.sh
docflow-studio
```

Los motores viven en `docflow-studio/backends/` y **siguen siendo CLIs
independientes**, utilizables desde la terminal como siempre:

| Backend                                               | Rol                                                                                                        | CLI               |
| ----------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- | ----------------- |
| [docflow](docflow-studio/backends/docflow/)           | Conversión documental masiva (→ Markdown, → PDF, Office ↔ ODF), watch, caché, reportes                     | `docflow`         |
| [pdf-suite](docflow-studio/backends/pdf-suite/)       | Manipulación PDF completa: comprimir, unir, dividir, rotar, OCR, cifrar, marca de agua, reparar, metadatos | `pdf-suite`       |
| [page-counter](docflow-studio/backends/page-counter/) | Conteo de páginas de los PDFs de la familia de blogs Quarto, con reporte Excel                             | `python3 main.py` |

Documentación: [docflow-studio/README.md](docflow-studio/README.md) ·
[docflow](docflow-studio/backends/docflow/README.md) ·
[pdf-suite](docflow-studio/backends/pdf-suite/README.md) ·
[page-counter](docflow-studio/backends/page-counter/README.md)

## Historia

- **v1** (`convert_ms_to_odf.sh`, 2024): script único MS Office → ODF.
- **v2** (`script_doc_suite`, 2025): suite modular con to-odf/to-pdf/to-md.
- **v3** (`docflow/`, 2026): reescritura completa como herramienta
  profesional basada en motores de conversión.
- **v4** (`docflow-studio/`, 2026): aplicación de escritorio unificada;
  docflow, pdf-suite y page-counter pasan a ser sus backends. Las versiones
  anteriores están en el historial de git.

## Licencia

MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
