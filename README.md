#  Conversor Masivo Microsoft Office → OpenDocument

![LibreOffice](https://img.shields.io/badge/LibreOffice-7%2B-blue) ![bash](https://img.shields.io/badge/bash-4%2B-green) ![license](https://img.shields.io/github/license/edisonachalma/convert_ms_to_odf.sh)

#readme

Script en Bash que convierte **recursivamente** todos tus archivos de Microsoft Office (.docx, .xlsx, .pptx, etc.) a formatos abiertos **OpenDocument** (.odt, .ods, .odp, …) usando LibreOffice en modo headless.

Mantiene la estructura exacta de carpetas y **elimina automáticamente los archivos originales** solo cuando la conversión ha sido exitosa.

Ideal para:

- Migrar toda una biblioteca de documentos a formatos libres
- Liberarte de la dependencia de Microsoft Office
- Preparar archivos para usar con OnlyOffice, LibreOffice, Collabora, etc.
- Hacer limpieza masiva sin miedo (¡solo borra lo que sí se convirtió bien!)

## Características

- Conversión 100 % recursiva
- Soporta todos los formatos habituales de Word, Excel y PowerPoint (incluidos los antiguos .doc, .xls, .ppt)
- Conserva la estructura de subcarpetas
- Solo elimina el archivo original si la conversión salió perfecta
- Colores y mensajes claros (verde = éxito, rojo = error, etc.)
- Resumen detallado al final
- Pregunta confirmación antes de empezar
- Muy seguro: los archivos que fallen se quedan intactos

## Requisitos

- **LibreOffice** instalado (el script usa `soffice`)
  ```bash
  # Debian / Ubuntu / Mint
  sudo apt update && sudo apt install libreoffice

  # Fedora
  sudo dnf install libreoffice

  # Arch / Manjaro
  sudo pacman -S libreoffice-fresh

  # macOS (con Homebrew)
  brew install --cask libreoffice
  ```

## Uso

```bash
# 1. Convierte todo lo que haya en la carpeta actual
./convert_ms_to_odf.sh

# 2. Convierte una carpeta concreta
./convert_ms_to_odf.sh "/ruta/a/mis/documentos"

# 3. Ver ayuda
./convert_ms_to_odf.sh --help
```

¡Eso es todo! El script te mostrará cuántos archivos encontró, pedirá confirmación y comenzará la conversión.

## Ejemplo de salida

```
[INFO] Directorio de trabajo: /home/yo/Documentos/ViejosOffice
[INFO] LibreOffice encontrado: LibreOffice 7.6

=== RESUMEN DE ARCHIVOS ENCONTRADOS ===

Documentos de texto:
 • .docx: 145 archivo(s)
 • .doc: 23 archivo(s)

Hojas de cálculo:
 • .xlsx: 67 archivo(s)

Presentaciones:
 • .pptx: 12 archivo(s)

[INFO] Total de archivos a procesar: 247

¿Desea continuar con la conversión? (s/N): s

# ... proceso con barra de progreso visual ...
# Al final:

[INFO] PROCESO COMPLETADO EXITOSAMENTE
[INFO] Eliminados: 247 archivo(s) originales
```

## Formatos soportados


| Tipo             | Entrada                                | Salida     |
| ---------------- | -------------------------------------- | ---------- |
| Documentos       | .docx, .doc, .dotx                     | .odt, .ott |
| Hojas de cálculo | .xlsx, .xls, .xlsm, .xltx              | .ods, .ots |
| Presentaciones   | .pptx, .ppt, .pptm, .ppsx, .pps, .potx | .odp, .otp |


## Licencia

**MIT License** – úsalo, modifícalo y distribúyelo libremente.

## Autor

Edison Achalma – 2024–2025  
¡Con mucho cariño para la comunidad de software libre!

---

**¡Dale una estrella si te ha salvado la vida migrando cientos de archivos!**
