# doc2md — Conversor recursivo Office → Markdown
#readme 

Convierte archivos `.docx`, `.odt`, `.pptx` y `.odp` a Markdown **en el mismo lugar** donde están los archivos originales, de forma recursiva.

---

## Cómo funciona

El `.md` se guarda **en el mismo directorio** que el archivo original y recibe el **mismo nombre**, solo cambia la extensión. Las imágenes se extraen ordenadas en una carpeta al lado del `.md`.

```
antes:
  ideas/2017-analisis/index.odt
  ideas/informe.docx

después:
  ideas/2017-analisis/index.odt          ← original intacto
  ideas/2017-analisis/index.md           ← mismo nombre
  ideas/2017-analisis/index_files/
      figure-md/
          fig-001.png
          fig-002.png

  ideas/informe.docx                     ← original intacto
  ideas/informe.md
  ideas/informe_files/
      figure-md/
          fig-001.png
```

---

## Instalación

### Ubicación del script

El script se aloja **siempre** en:

```
~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh
```

Para llamarlo desde cualquier lugar, crea un alias en tu `~/.zshrc` o `~/.bashrc`:

```bash
# En ~/.zshrc
alias doc2md="~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh"
```

O instálalo en el PATH:

```bash
chmod +x doc2md.sh
sudo install -m 755 doc2md.sh /usr/local/bin/doc2md
```

### Dependencias en Arch Linux

```bash
# Requeridas
sudo pacman -S pandoc python

# Opcionales (recomendadas)
sudo pacman -S libreoffice-still   # fallback para pptx/odp
sudo pacman -S imagemagick         # optimización de imágenes
```

---

## Uso

```
./doc2md.sh [OPCIONES] <DIRECTORIO>
```

El directorio es el único argumento posicional. Las opciones van antes.

### Opciones

| Opción                    | Descripción                                                                              |
| ------------------------- | ---------------------------------------------------------------------------------------- |
| `-f, --formats <lista>`   | Formatos a convertir, separados por coma. Por defecto: `docx,odt,pptx,odp`               |
| `-w, --overwrite`         | Sobreescribir `.md` ya existentes (por defecto se omiten)                                |
| `--img-format <fmt>`      | Formato de imágenes extraídas: `png` / `jpg` / `webp`. Por defecto: `png`                |
| `--img-quality <n>`       | Calidad para jpg/webp (1-100). Por defecto: `90`                                         |
| `--pandoc-args <args>`    | Argumentos extra para pandoc (entre comillas)                                            |
| `-e, --exclude <nombre>`  | Excluir carpeta por nombre o ruta absoluta. Usar varias veces para múltiples exclusiones |
| `--no-default-excludes`   | No aplicar la lista de exclusiones por defecto                                           |
| `--delete-converted`      | Eliminar los archivos Office **solo** si se convirtieron con éxito en esta ejecución     |
| `-l, --log <archivo>`     | Guardar log en archivo además de stdout                                                  |
| `-s, --summary <archivo>` | Generar archivo resumen de la conversión                                                 |
| `-v, --verbose`           | Salida detallada                                                                         |
| `-n, --dry-run`           | Simular sin escribir nada                                                                |
| `-h, --help`              | Mostrar ayuda                                                                            |

---

## Ejemplos

### Conversión básica de toda la carpeta ideas

```bash
./doc2md.sh ~/Documents/ideas
```

### Solo .odt y .docx, sobreescribir existentes

```bash
./doc2md.sh -f odt,docx -w ~/Documents/ideas
```

### Excluir carpetas específicas

```bash
# Excluir por nombre (afecta a todos los niveles)
./doc2md.sh -e notas -e borradores ~/Documents/ideas

# Excluir por lista separada por coma
./doc2md.sh -e "notas,borradores,trash" ~/Documents/ideas

# Excluir por ruta absoluta
./doc2md.sh -e /home/achalmaedison/Documents/ideas/privado ~/Documents/ideas
```

### Ver qué se convertiría sin hacer nada (dry-run)

```bash
./doc2md.sh -n -v ~/Documents/ideas
```

### Convertir y eliminar los originales convertidos

```bash
# Primero verificar con dry-run
./doc2md.sh -n --delete-converted ~/Documents/ideas

# Si todo se ve bien, ejecutar de verdad
./doc2md.sh --delete-converted ~/Documents/ideas
```

> **Nota:** `--delete-converted` solo elimina los archivos que fueron convertidos exitosamente **en esa misma ejecución**. No toca archivos que ya tenían `.md`, ni archivos que fallaron, ni ningún otro tipo de archivo.

### Con log y resumen

```bash
./doc2md.sh \
  -l ~/logs/doc2md.log \
  -s ~/logs/resumen.txt \
  -v \
  ~/Documents/ideas
```

### Imágenes en JPEG de alta calidad

```bash
./doc2md.sh --img-format jpg --img-quality 95 ~/Documents/ideas
```

### Pandoc con tabla de contenidos

```bash
./doc2md.sh --pandoc-args '--toc --toc-depth=3' ~/Documents/ideas
```

---

## Carpetas excluidas por defecto

El script excluye automáticamente estas carpetas en cualquier nivel:

```
.git   .svn   node_modules   __pycache__   .trash   Trash   .Trash
```

Para **agregar exclusiones permanentes**, edita la sección `EXCLUDE_DIRS` al inicio del script:

```bash
declare -a EXCLUDE_DIRS=(
    ".git"
    ".svn"
    "node_modules"
    "__pycache__"
    ".trash"
    "Trash"
    ".Trash"
    "privado"          # <- agrega las tuyas aquí
    "borradores"
)
```

Para **desactivar todas las exclusiones por defecto** en una ejecución:

```bash
./doc2md.sh --no-default-excludes ~/Documents/ideas
```

---

## Integración con tu flujo de trabajo

### Alias en zsh (recomendado)

Añade en `~/.zshrc`:

```bash
SCRIPTS_DIR="$HOME/Documents/scripts_for_libreoffice/script_convert_doc_to_md"
alias doc2md="$SCRIPTS_DIR/doc2md.sh"

# Atajo rápido para convertir el directorio actual
alias doc2md-here="$SCRIPTS_DIR/doc2md.sh ."
```

### Tarea cron (cada noche a las 02:00)

```bash
crontab -e
# Añadir:
0 2 * * * ~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh -w ~/Documents/ideas -l ~/logs/doc2md.log
```

### Función zsh con exclusiones frecuentes

```bash
# En ~/.zshrc
doc2md-ideas() {
  ~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh \
    -e "trash" -e "privado" \
    -w \
    "${1:-$HOME/Documents/ideas}"
}
```

---

## Estructura de salida detallada

Para un archivo `index.odt` con 2 imágenes:

```
<directorio>/
├── index.odt                    ← original intacto (o eliminado si usaste --delete-converted)
├── index.md                     ← Markdown generado
└── index_files/
    └── figure-md/
        ├── fig-001.png          ← primera imagen extraída
        └── fig-002.png          ← segunda imagen extraída
```

Si el documento no tiene imágenes, la carpeta `index_files/` no se crea.

---

## Dependencias

| Herramienta   | Rol                                 | Requerida      |
| ------------- | ----------------------------------- | -------------- |
| `pandoc`      | Motor principal de conversión       | ✅ Sí          |
| `python3`     | Corrección de rutas de imágenes     | ✅ Sí          |
| `libreoffice` | Fallback para `.pptx` y `.odp`      | ⚠️ Recomendada |
| `imagemagick` | Conversión/optimización de imágenes | ⭕ Opcional    |

---

## Solución de problemas

**Las presentaciones no se convierten bien**

```bash
sudo pacman -S libreoffice-still
```

**El script no tiene permisos de ejecución**

```bash
chmod +x ~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh
```

**Quiero ver exactamente qué va a hacer antes de ejecutar**

```bash
./doc2md.sh -n -v ~/Documents/ideas
```

**Un archivo genera un .md vacío o sin imágenes**
Ejecuta con `-v` para ver el detalle. Si el documento tiene imágenes vinculadas (no incrustadas), pandoc no puede extraerlas.

---

## Licencia

MIT © Edison Achalma — [github.com/achalmed](https://github.com/achalmed)
