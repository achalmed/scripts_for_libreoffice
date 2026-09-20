---
tipo: readme
estado: activo
---
# page-counter/ — cuenta las páginas de los index.pdf renderizados por cada blog Quarto y produce un reporte Excel (v2.0)

<!-- suite:inicio -->
**Suite `page_counter`** · objetivo *documentos* · estado *activo* · python · interfaz cli

Cuenta las páginas de los index.pdf generados por cada blog (APA) y produce un reporte por blog.

- Escribe en: ninguno · simula por defecto: sí
- Entrada: 04 index/_site y _pubs/*/_site
- Depende de: python3, pypdf

Comandos:

```bash
main.py
main.py --hub «04 index»
```

<sub>Bloque generado desde `suite.yml` por `core/suites.py generar` (2026-09-20); no se edita a mano.</sub>
<!-- suite:fin -->

> **Ubicación de los blogs (desde 2026-09-06):** los 11 `pub_*` son submódulos git de `website-achalma` y viven en `~/Documents/04 index/_pubs/pub_*`; el hub sigue en `~/Documents/04 index`. Las herramientas los localizan por esa subcarpeta (variable `SUBDIR_PUBS en config.py`), y aceptan el nombre de carpeta o el nombre corto sin `pub_`.

> Cuenta las páginas de los PDFs renderizados por la familia de blogs Quarto
> (`pub_*` y `website-achalma`) y genera un reporte Excel con subtotales por
> blog y total general.

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Arquitectura](#-arquitectura)
- [Bugs Corregidos](#-bugs-corregidos)
- [Solución de Problemas](#-solución-de-problemas)
- [Cómo Contribuir](#-cómo-contribuir)
- [Notas y Advertencias](#-notas-y-advertencias)

## 📖 Descripción

Recorre los `_site/` de cada blog buscando `index.pdf` (los PDF APA
generados por apaquarto) — o todos los PDFs con `--todos` — cuenta sus
páginas con `pypdf` y produce un `.xlsx` en `excel_databases/` con:

- **Hoja "Conteo de Páginas"**: bloques por blog con cada archivo, su número
  de páginas y estado (`OK` / `VACÍO` / `ERROR`), subtotales por blog y
  total general.
- **Hoja "Información"**: metadatos del reporte (fecha, tipo de búsqueda,
  totales).

Los blogs se seleccionan por su nombre lógico (`axiomata`, `chaska`, …); la
herramienta resuelve sola la carpeta real (`~/Documents/04 index/_pubs/pub_axiomata/_site`).
Las secciones de website-achalma se llaman `blog` y `teching` (o el alias
`website-achalma` para ambas).

## ⚙️ Requisitos

### Sistema Operativo

- Linux/macOS con Python >= 3.8.

### Dependencias

- `pypdf` >= 3.0 — lectura de PDFs (se acepta `PyPDF2` como fallback)
- `openpyxl` >= 3.0 — generación del Excel

## 🚀 Instalación

```bash
cd script_pdf_page_counter

# Con pip
pip install -r requirements.txt

# O con conda
conda install -c conda-forge pypdf openpyxl

# O el instalador asistido (crea entorno conda si hace falta)
./install.sh
```

## 💻 Uso

### Sintaxis

```bash
python3 main.py [OPCIONES]
```

### Opciones disponibles

| Flag              | Descripción                                          | Requerido |
| ----------------- | ---------------------------------------------------- | --------- |
| `-b, --blogs ...` | Blogs específicos (separados por espacios)           | No        |
| `-t, --todos`     | Buscar todos los PDFs, no solo `index.pdf`           | No        |
| `-o, --output F`  | Nombre del Excel de salida (en `excel_databases/`)   | No        |
| `-l, --listar`    | Listar blogs disponibles y salir                     | No        |
| `-v, --verbose`   | Detalle extra (incluye la causa de PDFs ilegibles)   | No        |
| `--version`       | Mostrar versión                                      | No        |
| `-h, --help`      | Mostrar ayuda                                        | No        |

### Ejemplos de uso

```bash
# Todos los blogs, solo index.pdf, nombre con timestamp
python3 main.py

# Blogs específicos
python3 main.py -b actus-mercator aequilibria

# Todos los PDFs (no solo index.pdf) con salida personalizada
python3 main.py --todos -o conteo_completo.xlsx

# Ver qué blogs están renderizados y disponibles
python3 main.py --listar
```

Más ejemplos y casos de uso en [ejemplos.md](ejemplos.md).

## 🗂️ Arquitectura

```
script_pdf_page_counter/
├── main.py              # Punto de entrada — solo orquestación
├── config.py            # Rutas de blogs, prefijo pub_, códigos de salida
├── requirements.txt     # Dependencias pip
├── install.sh           # Instalador asistido (conda/pip)
└── lib/
    ├── __init__.py
    ├── logger.py        # Logger único de la aplicación (--verbose → DEBUG)
    ├── cli.py           # Parser argparse centralizado
    ├── validator.py     # Ruta base, nombres de blog, resolución de rutas
    ├── scanner.py       # Búsqueda de PDFs y conteo de páginas
    ├── excel_report.py  # Construcción del .xlsx (datos + metadatos)
    └── ui.py            # Encabezados, secciones, listado y resumen
```

### Descripción de módulos

| Archivo               | Responsabilidad                                         |
| --------------------- | ------------------------------------------------------- |
| `main.py`             | Orquestar: validar → escanear → reportar → resumen      |
| `config.py`           | Única fuente de rutas y constantes editables            |
| `lib/logger.py`       | Formato de log consistente en toda la app               |
| `lib/cli.py`          | Definición completa de la CLI                           |
| `lib/validator.py`    | Fallar temprano ante rutas o nombres inválidos          |
| `lib/scanner.py`      | Única pieza que abre PDFs                               |
| `lib/excel_report.py` | Única pieza que escribe Excel                           |
| `lib/ui.py`           | Presentación en terminal (no toca PDFs ni Excel)        |

## 🐛 Bugs Corregidos

### Bug #1: Ruta base apuntando a una carpeta inexistente
- **Descripción**: `RUTA_BASE_PUBLICACIONES` era
  ~/Documents/publicaciones, una estructura que ya no existe — los blogs
  viven como submódulos `pub_<nombre>` del hub, en `~/Documents/04 index/_pubs/` (`SUBDIR_PUBS` en `config.py`).
- **Impacto**: la herramienta terminaba siempre con "La ruta base no
  existe"; **ninguna** funcionalidad era utilizable.
- **Corrección**: `config.py` apunta a `~/Documents` con prefijo `pub_`;
  `lib/validator.py` resuelve nombre lógico → carpeta real. Verificado
  contra los 13 blogs reales (11 `pub_*` + `blog` + `teching`).

### Bug #2: Errores de lectura de PDF silenciados
- **Descripción**: `except Exception as e: return -1` descartaba la causa
  (el `e` ni se usaba).
- **Impacto**: un PDF corrupto aparecía como "ERROR" sin forma de saber por
  qué (cifrado, truncado, permisos…).
- **Corrección**: la causa se registra en nivel DEBUG; `--verbose` la
  muestra por archivo.

### Bug #3: Sin códigos de salida
- **Descripción**: todos los caminos de error hacían `return` desde `main()`
  y el proceso terminaba con 0.
- **Impacto**: imposible detectar fallos desde scripts o cron
  (`main.py && siguiente` seguía adelante tras un error).
- **Corrección**: códigos estándar — 0 éxito, 1 error (incluye PDFs
  ilegibles), 2 uso, 3 ruta/blogs no encontrados, 5 dependencia faltante.

### Bug #4: Typos en nombres de blog ignorados en silencio
- **Descripción**: `-b axiomta` (typo) simplemente no matcheaba y el blog
  quedaba fuera del reporte sin aviso.
- **Impacto**: reportes incompletos sin que el usuario lo notara.
- **Corrección**: los nombres desconocidos abortan con error explícito y
  sugerencia de `--listar` (salida 2).

### Bug #5: Archivos vacíos contados como exitosos
- **Descripción**: el resumen calculaba "exitosos = total − errores", con lo
  que los PDFs de 0 páginas (estado `VACÍO`) contaban como éxito.
- **Impacto**: resumen final engañoso.
- **Corrección**: los tres estados se cuentan por separado y el resumen
  muestra una línea propia para vacíos.

### Bug #6: Dependencia PDF exigida para todo
- **Descripción**: el import de la librería PDF a nivel de módulo abortaba
  incluso `--listar` y `--help`, que no leen ningún PDF.
- **Impacto**: la herramienta era inutilizable sin la librería aunque solo
  se quisiera consultar los blogs disponibles.
- **Corrección**: import diferido — la dependencia solo se exige al empezar
  a escanear. Además se migró a `pypdf` (PyPDF2 está descontinuado desde
  2023), manteniendo PyPDF2 como fallback.

### Bug #7: Fallo al guardar el Excel sin capturar
- **Descripción**: `wb.save()` sin manejo de errores.
- **Impacto**: un destino sin permisos o el archivo abierto en LibreOffice
  producía un traceback crudo tras minutos de escaneo.
- **Corrección**: `PermissionError`/`OSError` capturados con mensaje claro y
  salida 1.

## 🔧 Solución de Problemas

### Error: "no está instalado pypdf"

```bash
pip install -r requirements.txt
```

### "No se encontraron blogs para procesar"

Los `_site/` no existen aún: renderiza los blogs con Quarto
(`quarto render`) o verifica cuáles están disponibles con `--listar`.

### El Excel no se guarda (Permission denied)

Cierra el archivo si lo tienes abierto en LibreOffice/Excel y reintenta.

### Un archivo aparece como ERROR

Ejecuta con `-v` para ver la causa exacta (PDF cifrado, truncado, etc.).

## 🤝 Cómo Contribuir

1. Crea el módulo en `lib/<nuevo_modulo>.py` con una única responsabilidad.
2. Añade sus flags en `lib/cli.py` y sus constantes en `config.py`.
3. Impórtalo desde `main.py`; mantén `main()` como orquestador puro.
4. Verifica con `python3 -m py_compile main.py config.py lib/*.py`.

### Estándares de código

- Máximo ~30 líneas por función; docstrings que expliquen el "por qué".
- Nada de `except Exception: pass` — captura tipos específicos o registra
  la causa.
- Rutas y constantes solo en `config.py`.

## ⚠️ Notas y Advertencias

- El conteo opera sobre los PDFs **renderizados** en `_site/`; con
  `freeze: true` en Quarto un render global puede reutilizar caché — si un
  conteo parece desactualizado, re-renderiza el blog afectado.
- Los nombres lógicos de blog no llevan el prefijo `pub_` (se usa
  `axiomata`, no `pub_axiomata`); el prefijo es configurable en `config.py`.
- El estado `VACÍO` (0 páginas) es raro pero posible en PDFs malformados
  que pypdf sí puede abrir.
- `excel_databases/` local a esta herramienta (ignorado en git) no es el Excel de
  metadatos de los blogs, que vive en
  `scripts_quarto_studio/backend/script_metadata_manager/excel_databases/`;
  son almacenes distintos.

## Límite honesto

- **Cuenta lo renderizado, no lo escrito**: opera sobre los `index.pdf` de `_site/`; con `freeze: true` un
  conteo viejo significa un render viejo, no un bug.
- **No escribe fuera de su carpeta**: solo el Excel en `excel_databases/`; por eso simula por defecto en el
  contrato de suites (`escribe_en: ninguno`).
- **La raíz de los blogs se resuelve con `Path.home()`** en `config.py` (`DIR_HUB = "04 index"`,
  `SUBDIR_PUBS = "04 index/_pubs"`), no con `core/env.py`: si el workspace no está en `~/Documents`, hay que
  editar `config.py` o pasar `--hub`.
- Los nombres lógicos de blog no llevan `pub_`; `blog` y `teching` cuelgan del `_site/` del hub y no tienen
  `_site/` propio.
- Sin pruebas automáticas; `pypdf` es la única dependencia externa y sin ella no arranca.
