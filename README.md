# scripts_for_libreoffice

Repositorio de herramientas para la conversión y gestión de documentos
Office/LibreOffice en Linux.

## docflow — suite de conversión documental

El proyecto principal de este repositorio es **[docflow](docflow/)**, una
herramienta CLI profesional para la conversión masiva de documentos:

- **→ Markdown** (Obsidian, Logseq, MkDocs, Hugo, Quarto, GitHub): imágenes
  extraídas y ordenadas, notas del presentador, tablas, metadatos como
  frontmatter YAML.
- **→ PDF** con el mejor motor disponible, más operaciones completas sobre
  PDFs: unir, dividir, comprimir, PDF/A, cifrar, marca de agua, OCR.
- **Office ↔ OpenDocument** bidireccional, incluidas plantillas.

```bash
cd docflow && ./install.sh
docflow to-md ~/Documentos/tesis/
```

Documentación completa: [docflow/README.md](docflow/README.md) ·
Arquitectura: [docflow/docs/ARCHITECTURE.md](docflow/docs/ARCHITECTURE.md)

## Historia

- **v1** (`convert_ms_to_odf.sh`, 2024): script único MS Office → ODF.
- **v2** (`script_doc_suite`, 2025): suite modular con to-odf/to-pdf/to-md.
- **v3** (`docflow/`, 2026): reescritura completa como herramienta
  profesional basada en motores de conversión. Las versiones anteriores
  están disponibles en el historial de git.

## Licencia

MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
