---
tipo: doc
titulo: "Visión de producto — DocFlow Studio (2026-07-13)"
estado: hecho
creado: 2026-07-13
---
Este proyecto me parece incluso **más prometedor** que *Quarto Studio*. La razón es que no está limitado a un ecosistema (Quarto), sino que puede convertirse en una plataforma general para el procesamiento documental. Si lo diseñas bien, podría servir tanto a investigadores como a empresas, universidades o incluso editoriales.

Sin embargo, creo que todavía lo estás viendo como una **colección de herramientas** (`docflow`, `pdf-suite`, `page-counter`) con una interfaz gráfica encima. Yo lo replantearía como un **Document Processing Platform**.

---

# Mi visión

> Visión de producto escrita en el vault el 2026-07-13 (`01 notes/proyecto-document-studio.md`, tipo idea) y trasladada a `docs/` del repo en
> DOC8 (2026-09-20, decisión D15): es el documento fundacional de lo que hoy es DocFlow Studio. Lo construido y vigente está en
> `README.md` y `CLAUDE.md`; lo que aquí se prometió y no se hizo es historia, no pendiente.

No lo llamaría únicamente **Document Studio**.

Lo pensaría como algo así:

> **Document Studio**
> 
> *An Integrated Development Environment for Documents*

Así como:

- VS Code → Código

- RStudio → R

- Quarto Studio → Quarto

- Document Studio → Documentos

Es decir, un IDE completo para documentos.

---

# Lo primero que cambiaría

Actualmente tienes

```text
backends/

docflow

pdf-suite

page-counter
```

Eso está bien, pero yo agregaría un nivel más.

```text
backends/

engines/

docflow

pdf-suite

page-counter

plugins/

ocr

latex

pandoc

metadata

templates

watcher
```

Los *engines* hacen el trabajo pesado.

Los *plugins* agregan funcionalidades.

---

# El núcleo del programa

Yo crearía un verdadero núcleo.

```text
studio/

core/

project.py

document.py

workspace.py

plugin_manager.py

task_manager.py

event_bus.py

command_bus.py
```

Ese sería el corazón.

Todo hablaría con él.

---

# El concepto de Workspace

Actualmente parece trabajar sobre archivos.

Yo trabajaría sobre proyectos.

Ejemplo

```text
Investigacion/

tesis/

libros/

papers/

diapositivas/

imagenes/

bibliografia/
```

El usuario abre un Workspace.

No un archivo.

---

# Document Model

Tendrías una clase

```python
Document
```

que conozca

```text
nombre

ruta

tipo

paginas

autor

titulo

metadata

idioma

formato

checksum

estado
```

Entonces cualquier módulo trabaja igual.

No importa si es

```
pdf

docx

odt

pptx

xlsx
```

---

# Pipeline de procesamiento

En lugar de ejecutar scripts directamente.

Haría pipelines.

```text
DOCX

↓

Metadata

↓

OCR

↓

Markdown

↓

Normalize

↓

Validate

↓

Report

↓

Export
```

Cada paso puede activarse o no.

---

# Procesamiento masivo

Aquí creo que está el mayor potencial.

Ejemplo

Selecciono

```
2000 archivos
```

Y digo

↓

Convertir

↓

Markdown

↓

Conservar estructura

↓

Extraer imágenes

↓

OCR

↓

Generar índice

↓

Crear reporte

↓

Listo.

Todo en paralelo.

---

# Explorador inteligente

No un árbol de archivos.

Sino un navegador.

```
Todos los PDF

Todos los DOCX

Todos los PPTX

Todos los archivos dañados

Todos los archivos sin metadatos

Todos los documentos grandes

Duplicados

Sin OCR
```

---

# Dashboard

Mostraría

```
Workspace

1245 documentos

352 PDF

98 DOCX

430 ODT

132 PPTX

Errores

12

Duplicados

33

OCR pendiente

44

Conversión pendiente

15
```

---

# Metadata Manager

Aquí invertiría mucho esfuerzo.

Cada documento tendría

```
Autor

Título

Idioma

Tema

Fecha

Palabras clave

Resumen

Editorial

Institución
```

Y todo editable desde la GUI.

---

# Motor de búsqueda

Algo tipo

```
Buscar

autor:economía

pages>100

type:pdf

date>2024

keyword:panel
```

Como Obsidian.

---

# Sistema de etiquetas

Cada documento

```
tesis

economía

microeconomía

panel

ayacucho
```

Y luego navegar por etiquetas.

---

# OCR integrado

No solamente OCR.

Elegir motor.

```
Tesseract

PaddleOCR

OCRmyPDF

EasyOCR
```

Y comparar resultados.

---

# IA

Aquí Document Studio puede ser impresionante.

Por ejemplo

Selecciono un PDF.

↓

Botón

```
Analizar con IA
```

Obtengo

```
Resumen

Palabras clave

Título sugerido

Clasificación

Idioma

Nivel académico

Citas

Bibliografía
```

---

# Sistema de conversiones

En vez de

```
Convertir

↓

PDF
```

Mostrar un grafo.

```
DOCX

↓

Markdown

↓

LaTeX

↓

PDF

↓

HTML

↓

EPUB

↓

DOCX
```

El usuario ve todas las rutas posibles.

---

# Integración con LaTeX

Aquí tienes una ventaja enorme por tu experiencia.

Podría tener

```
Compilar

BibTeX

Biber

Latexmk

TikZ

Glossaries

MakeIndex
```

Desde la GUI.

---

# Plantillas

Una sección completa.

```
Artículo APA

Libro

Informe

Tesis

Presentación

Reporte

Curriculum

Carta
```

---

# Watch Mode

Ya tienes algo parecido.

Yo lo haría como VS Code.

```
Cambia el archivo

↓

Reconvierte

↓

Actualiza índice

↓

Actualiza preview

↓

Genera reporte
```

---

# Preview

Importantísimo.

Panel derecho.

```
Markdown

↓

Render
```

```
PDF

↓

Preview
```

```
DOCX

↓

Preview
```

Todo sin salir.

---

# Sistema de Reportes

No solo Excel.

También

```
CSV

Markdown

PDF

HTML

JSON
```

---

# API interna

Todo debería ser invocable.

Ejemplo

```python
Studio.convert()

Studio.metadata()

Studio.pdf.compress()

Studio.watch()

Studio.ocr()

Studio.report()
```

---

# Arquitectura

Yo reorganizaría algo así:

```text
Document Studio
│
├── Core
│   ├── Workspace Manager
│   ├── Plugin Manager
│   ├── Document Model
│   ├── Task Manager
│   ├── Event Bus
│   ├── Command Bus
│   └── Settings
│
├── UI
│   ├── Dashboard
│   ├── Explorer
│   ├── Search
│   ├── Metadata
│   ├── Preview
│   ├── Reports
│   ├── Console
│   └── Tasks
│
├── Services
│   ├── Conversion
│   ├── OCR
│   ├── Metadata
│   ├── PDF
│   ├── Office
│   ├── Markdown
│   ├── LaTeX
│   ├── Reports
│   └── AI
│
├── Engines
│   ├── Pandoc
│   ├── LibreOffice
│   ├── OCRmyPDF
│   ├── Tesseract
│   ├── Ghostscript
│   ├── Poppler
│   └── ImageMagick
│
├── Plugins
│   ├── Templates
│   ├── Academic
│   ├── Publishing
│   ├── Translation
│   ├── Citation
│   └── Custom
│
└── Projects
    ├── Research
    ├── Books
    ├── Courses
    └── Archives
```

## Una oportunidad mayor: una suite integrada

Viendo este proyecto y el anterior, creo que ambos resuelven problemas complementarios. En lugar de mantener dos aplicaciones completamente separadas, podrías diseñar una **suite modular** donde compartan el mismo núcleo (`Core`, `Workspace`, `Task Manager`, `Plugin Manager`, `Console`, `Settings`) y se diferencien por sus módulos especializados.

Por ejemplo:

- **Core Studio**: infraestructura común (proyectos, tareas, configuración, plugins, eventos).

- **Document Studio**: conversión documental, OCR, PDF, Office, LaTeX.

- **Quarto Studio**: publicación, gestión de blogs, metadatos Quarto, índices, sitios web.

- En el futuro podrías añadir módulos como **Academic Studio** (tesis, artículos científicos y APA), **Data Studio** (datasets y estadísticas) o **Course Studio** (materiales docentes y exámenes).

Con esa estrategia evitarías duplicar código, tendrías una experiencia de usuario consistente y podrías evolucionar hacia una plataforma mucho más grande y mantenible, donde cada estudio es simplemente un conjunto de plugins sobre un mismo núcleo. Esa arquitectura es la que mejor escala si tu objetivo es convertir estos proyectos en herramientas de largo plazo para investigación, docencia y publicación académica.
