# doc2md — Conversor recursivo Office → Markdown

**Convierte archivos `.docx`, `.odt`, `.pptx` y `.odp` a Markdown de forma recursiva**, replicando la estructura de directorios y extrayendo imágenes ordenadas por número de figura.

---

## Características

- 🔁 **Recursivo** — escanea toda la jerarquía de carpetas
- 📁 **Estructura limpia** — cada documento genera su propio `index.md` + carpeta de figuras
- 🖼️ **Imágenes organizadas** — extraídas y renombradas como `fig-001.png`, `fig-002.png`…
- ⚡ **Múltiples formatos** — `docx`, `odt`, `pptx`, `odp`
- 🔄 **Fallback LibreOffice** — para presentaciones cuando pandoc falla
- 🛡️ **Seguro** — no sobreescribe por defecto; usa `-w` para forzar
- 📋 **Log y resumen** — registra todo en archivo y genera informe final
- 🏃 **Dry-run** — simula sin escribir nada (`-n`)

---

## Estructura de salida

Por cada archivo convertido:

```
<salida>/
└── <subcarpeta_original>/
    └── <nombre_documento>/
        ├── index.md                  ← Contenido en Markdown
        └── index_files/
            └── figure-md/
                ├── fig-001.png       ← Imágenes extraídas, ordenadas
                ├── fig-002.png
                └── ...
```

### Ejemplo real

```
~/Documentos/
├── informe_2024.docx
├── presentacion.pptx
└── trabajos/
    └── tesis.odt
```

Se convierte en:

```
~/Markdown/
├── informe_2024/
│   ├── index.md
│   └── index_files/figure-md/
│       ├── fig-001.png
│       └── fig-002.png
├── presentacion/
│   ├── index.md
│   └── index_files/figure-md/
│       └── fig-001.png
└── trabajos/
    └── tesis/
        └── index.md
```

---

## Instalación

### 1. Clonar o copiar el script

```bash
# Opción A — directamente
cp doc2md.sh ~/bin/doc2md
chmod +x ~/bin/doc2md

# Opción B — en /usr/local/bin (disponible para todos los usuarios)
sudo install -m 755 doc2md.sh /usr/local/bin/doc2md
```

### 2. Dependencias en Arch Linux

```bash
# Requeridas
sudo pacman -S pandoc python

# Opcionales (recomendadas)
sudo pacman -S libreoffice-still   # fallback para pptx/odp
sudo pacman -S imagemagick         # optimización de imágenes
```

### 3. Verificar instalación

```bash
./doc2md.sh --help
```

---

## Uso

```
./doc2md.sh [OPCIONES] -i <DIRECTORIO_ENTRADA> -o <DIRECTORIO_SALIDA>
```

### Opciones

| Opción                    | Descripción                             | Por defecto         |
| ------------------------- | --------------------------------------- | ------------------- |
| `-i, --input <dir>`       | Directorio fuente (requerido)           | —                   |
| `-o, --output <dir>`      | Directorio de salida (requerido)        | —                   |
| `-f, --formats <lista>`   | Formatos separados por coma             | `docx,odt,pptx,odp` |
| `-w, --overwrite`         | Sobreescribir `.md` existentes          | No                  |
| `--flatten`               | No replicar subdirectorios              | No                  |
| `--img-format <fmt>`      | Formato de imágenes: `png`/`jpg`/`webp` | `png`               |
| `--img-quality <n>`       | Calidad jpg/webp (1-100)                | `90`                |
| `--pandoc-args <args>`    | Argumentos extra para pandoc            | —                   |
| `-l, --log <archivo>`     | Guardar log en archivo                  | —                   |
| `-s, --summary <archivo>` | Generar resumen de conversión           | —                   |
| `-v, --verbose`           | Salida detallada                        | No                  |
| `-n, --dry-run`           | Simular sin escribir                    | No                  |
| `-h, --help`              | Mostrar ayuda                           | —                   |

---

## Ejemplos

### Conversión básica

```bash
./doc2md.sh -i ~/Documentos -o ~/Markdown
```

### Solo Word y ODT, sobreescribir existentes

```bash
doc2md -i ~/Documentos -o ~/Markdown -f docx,odt -w
```

### Con log detallado y resumen

```bash
doc2md -i ~/Documentos -o ~/Markdown \
       -l ~/doc2md.log \
       -s ~/doc2md-resumen.txt \
       -v
```

### Imágenes en JPEG de alta calidad

```bash
doc2md -i ~/Documentos -o ~/Markdown \
       --img-format jpg \
       --img-quality 95
```

### Dry-run para previsualizar

```bash
doc2md -i ~/Documentos -o ~/Markdown -n -v
```

### Aplanar estructura (todo al mismo nivel)

```bash
doc2md -i ~/Documentos -o ~/Markdown --flatten
```

### Pandoc con tabla de contenidos

```bash
doc2md -i ~/Documentos -o ~/Markdown --pandoc-args '--toc --toc-depth=3'
```

---

## Dependencias detalladas

| Herramienta   | Rol                                 | Requerida      |
| ------------- | ----------------------------------- | -------------- |
| `pandoc`      | Motor principal de conversión       | ✅ Sí          |
| `python3`     | Corrección de rutas de imágenes     | ✅ Sí          |
| `libreoffice` | Fallback para `.pptx` y `.odp`      | ⚠️ Recomendada |
| `imagemagick` | Conversión/optimización de imágenes | ⭕ Opcional    |

---

## Notas sobre formatos

### DOCX / ODT

Pandoc convierte directamente con excelente fidelidad. Las imágenes incrustadas se extraen automáticamente.

### PPTX / ODP

Pandoc soporta estas conversiones; el script intenta pandoc primero y, si falla, usa LibreOffice para pre-convertir a ODT y luego aplica pandoc. Se recomienda tener LibreOffice instalado.

### Imágenes SVG

Se conservan como SVG cuando pandoc las extrae en ese formato. Para convertirlas a PNG/JPG, asegúrate de tener ImageMagick instalado.

---

## Solución de problemas

### "pandoc: Could not find image"

Asegúrate de ejecutar el script con la ruta **absoluta** al directorio fuente o desde el mismo directorio.

### Las presentaciones no se convierten bien

Instala LibreOffice: `sudo pacman -S libreoffice-still`

### Las imágenes no aparecen en el Markdown

Verifica que el archivo original tenga imágenes incrustadas (no vinculadas externamente). Usa `-v` para ver el detalle del proceso.

### Error "Permission denied"

```bash
chmod +x doc2md.sh
```

---

## Integración con sistemas de documentación

El formato de salida es compatible con:

- **Quarto** — los `index.md` se pueden renombrar a `index.qmd`
- **MkDocs** — estructura lista para `docs/`
- **Jekyll / Hugo** — añade frontmatter con `--pandoc-args '--standalone'`
- **Obsidian** — copia la carpeta de salida como vault

---

## Automatización con cron / systemd

### Tarea cron (cada día a las 02:00)

```bash
crontab -e
# Añadir:
0 2 * * * /usr/local/bin/doc2md -i ~/Documentos -o ~/Markdown -w -l ~/logs/doc2md.log
```

### Servicio systemd (oneshot)

```ini
# ~/.config/systemd/user/doc2md.service
[Unit]
Description=Conversión Office a Markdown

[Service]
Type=oneshot
ExecStart=/usr/local/bin/doc2md -i %h/Documentos -o %h/Markdown -w
```

```bash
systemctl --user enable --now doc2md.service
```

---

## Licencia

MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
