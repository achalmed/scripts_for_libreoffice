# Changelog

## v3.0.0 — 2026-07-10

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
