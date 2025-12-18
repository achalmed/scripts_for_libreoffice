#!/bin/bash

# Script para buscar recursivamente archivos .odt y .odp en el directorio actual
# y sus subcarpetas, y convertirlos a PDF en paralelo usando LibreOffice
# Los PDFs se guardan en la misma carpeta que el archivo original con el mismo nombre

# Verifica si LibreOffice está instalado
if ! command -v libreoffice &> /dev/null; then
    echo "Error: LibreOffice no está instalado. Instálalo con: sudo apt install libreoffice"
    exit 1
fi

# Verifica si parallel está instalado
if ! command -v parallel &> /dev/null; then
    echo "Error: parallel no está instalado. Instálalo con: sudo apt install parallel"
    exit 1
fi

# Configura LD_LIBRARY_PATH para evitar errores de libreglo.so
export LD_LIBRARY_PATH=/usr/lib/libreoffice/program:$LD_LIBRARY_PATH

# Función para convertir un archivo a PDF
convert_to_pdf() {
    local file="$1"
    local filename=$(basename "$file")
    local dirname=$(dirname "$file")
    local output_file="$dirname/${filename%.*}.pdf"

    # Verifica si el archivo PDF ya existe
    if [ -f "$output_file" ]; then
        echo "El archivo $output_file ya existe, omitiendo conversión."
        return 0
    fi

    echo "Convirtiendo: $file -> $output_file"
    libreoffice --headless --convert-to pdf --outdir "$dirname" "$file" 2>&1 | while IFS= read -r line; do
        echo "  $line"
    done

    if [ ${PIPESTATUS[0]} -eq 0 ]; then
        echo "Conversión exitosa: $output_file"
    else
        echo "Error al convertir: $file"
        return 1
    fi
}

# Exporta la función para que parallel pueda usarla
export -f convert_to_pdf

# Buscar y convertir archivos en el directorio actual y subcarpetas en paralelo
echo "Buscando archivos .odt y .odp en $(pwd)..."
find "$(pwd)" -type f \( -iname "*.odt" -o -iname "*.odp" \) -print0 | parallel -0 --jobs 4 convert_to_pdf {}

echo "Conversión completada. Los PDFs están en las mismas carpetas que los archivos originales."
