#!/bin/bash

################################################################################
# Script: convert_ms_to_odf.sh
# Descripción: Convierte recursivamente archivos de Microsoft Office a formatos 
#              OpenDocument manteniendo la estructura de directorios original.
#              Elimina los archivos originales después de conversión exitosa.
# Autor: Edison Achalma
# Fecha: 2024
# Requisitos: LibreOffice (soffice)
# Uso: ./convert_ms_to_odf.sh [directorio]
#      Si no se especifica directorio, usa el actual (.)
################################################################################

# Colores para mensajes
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

################################################################################
# CONFIGURACIÓN
################################################################################

# Directorio de trabajo (por defecto el actual)
WORK_DIR="${1:-.}"

# Array asociativo para rastrear conversiones exitosas
declare -A CONVERTED_FILES

################################################################################
# FUNCIONES DE LOGGING
################################################################################

log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[ADVERTENCIA]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

log_detail() {
    echo -e "${BLUE}  →${NC} $1"
}

################################################################################
# FUNCIONES DE VALIDACIÓN
################################################################################

# Verificar si el directorio existe
check_directory() {
    if [ ! -d "$WORK_DIR" ]; then
        log_error "El directorio '$WORK_DIR' no existe"
        exit 1
    fi
    
    # Convertir a ruta absoluta
    WORK_DIR=$(cd "$WORK_DIR" && pwd)
    log_info "Directorio de trabajo: $WORK_DIR"
}

# Verificar si LibreOffice está instalado
check_libreoffice() {
    if ! command -v soffice &> /dev/null; then
        log_error "LibreOffice no está instalado o no se encuentra en PATH"
        log_error "Instala LibreOffice: sudo apt install libreoffice (Debian/Ubuntu)"
        exit 1
    fi
    log_info "LibreOffice encontrado: $(soffice --version)"
}

################################################################################
# FUNCIONES DE CONTEO Y BÚSQUEDA
################################################################################

# Contar archivos por extensión en el directorio de trabajo
count_files() {
    local extension="$1"
    find "$WORK_DIR" -type f -name "*.$extension" 2>/dev/null | wc -l
}

# Obtener lista de todos los archivos a convertir
get_all_files() {
    local pattern="$1"
    find "$WORK_DIR" -type f -name "*.$pattern" -print0 2>/dev/null
}

################################################################################
# FUNCIÓN PRINCIPAL DE CONVERSIÓN
################################################################################

convert_files() {
    local pattern="$1"
    local output_format="$2"
    local description="$3"
    
    local count=$(count_files "$pattern")
    
    if [ "$count" -eq 0 ]; then
        log_warning "No se encontraron archivos .$pattern"
        return 0
    fi
    
    log_info "Procesando $count archivo(s) $description ($pattern → $output_format)..."
    
    local success=0
    local failed=0
    local skipped=0
    
    # Procesar cada archivo encontrado
    while IFS= read -r -d '' file; do
        local dir=$(dirname "$file")
        local base=$(basename "$file")
        local name="${base%.*}"
        local output_file="$dir/$name.$output_format"
        
        # Verificar si el archivo de salida ya existe
        if [ -f "$output_file" ]; then
            log_detail "⊘ Ya existe: $file → $(basename "$output_file")"
            ((skipped++))
            # Marcar como exitoso para permitir eliminación del original
            CONVERTED_FILES["$file"]="skipped"
            continue
        fi
        
        # Mostrar ruta relativa desde el directorio de trabajo
        local rel_path="${file#$WORK_DIR/}"
        log_detail "Convirtiendo: $rel_path"
        
        # Realizar la conversión
        if soffice --headless --convert-to "$output_format" --outdir "$dir" "$file" &>/dev/null; then
            ((success++))
            # Marcar archivo como convertido exitosamente
            CONVERTED_FILES["$file"]="success"
            log_detail "✓ Completado: $(basename "$output_file")"
        else
            ((failed++))
            log_error "✗ Falló la conversión de: $rel_path"
        fi
    done < <(get_all_files "$pattern")
    
    # Mostrar resumen
    if [ $skipped -gt 0 ]; then
        log_info "Resumen: $success exitosos, $failed fallidos, $skipped omitidos (ya existían)"
    else
        log_info "Resumen: $success exitosos, $failed fallidos"
    fi
    
    return $failed
}

################################################################################
# FUNCIÓN DE ELIMINACIÓN
################################################################################

delete_originals() {
    local pattern="$1"
    local deleted=0
    local kept=0
    
    log_info "Verificando archivos .$pattern para eliminar..."
    
    while IFS= read -r -d '' file; do
        # Solo eliminar si la conversión fue exitosa o si ya existía el archivo convertido
        if [[ "${CONVERTED_FILES[$file]}" == "success" ]] || [[ "${CONVERTED_FILES[$file]}" == "skipped" ]]; then
            local rel_path="${file#$WORK_DIR/}"
            rm "$file"
            ((deleted++))
            log_detail "🗑 Eliminado: $rel_path"
        else
            ((kept++))
        fi
    done < <(get_all_files "$pattern")
    
    if [ $deleted -gt 0 ]; then
        log_info "Eliminados: $deleted archivo(s) .$pattern"
    fi
    
    if [ $kept -gt 0 ]; then
        log_warning "Conservados: $kept archivo(s) .$pattern (conversión falló)"
    fi
}

################################################################################
# FUNCIÓN DE RESUMEN INICIAL
################################################################################

show_summary() {
    log_info "=== RESUMEN DE ARCHIVOS ENCONTRADOS ==="
    
    local total=0
    local extensions=("docx" "doc" "dotx" "xlsx" "xls" "xlsm" "xltx" "pptx" "ppt" "pptm" "ppsx" "pps" "potx")
    
    echo ""
    echo "Documentos de texto:"
    for ext in "docx" "doc" "dotx"; do
        local count=$(count_files "$ext")
        if [ $count -gt 0 ]; then
            echo "  • .$ext: $count archivo(s)"
            ((total+=count))
        fi
    done
    
    echo ""
    echo "Hojas de cálculo:"
    for ext in "xlsx" "xls" "xlsm" "xltx"; do
        local count=$(count_files "$ext")
        if [ $count -gt 0 ]; then
            echo "  • .$ext: $count archivo(s)"
            ((total+=count))
        fi
    done
    
    echo ""
    echo "Presentaciones:"
    for ext in "pptx" "ppt" "pptm" "ppsx" "pps" "potx"; do
        local count=$(count_files "$ext")
        if [ $count -gt 0 ]; then
            echo "  • .$ext: $count archivo(s)"
            ((total+=count))
        fi
    done
    
    echo ""
    log_info "Total de archivos a procesar: $total"
    echo ""
    
    if [ $total -eq 0 ]; then
        log_warning "No se encontraron archivos de Microsoft Office para convertir"
        exit 0
    fi
}

################################################################################
# SCRIPT PRINCIPAL
################################################################################

main() {
    echo ""
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    log_info "║  Conversión MS Office → OpenDocument Format                  ║"
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    echo ""
    
    # Validaciones iniciales
    check_directory
    check_libreoffice
    echo ""
    
    # Mostrar resumen de archivos encontrados
    show_summary
    
    # Confirmar antes de proceder
    read -p "¿Desea continuar con la conversión? (s/N): " confirm
    if [[ ! "$confirm" =~ ^[sS]$ ]]; then
        log_warning "Operación cancelada por el usuario"
        exit 0
    fi
    echo ""
    
    # Contador total de errores
    local total_errors=0
    
    # === DOCUMENTOS DE TEXTO ===
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    log_info "║  DOCUMENTOS DE TEXTO                                          ║"
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    convert_files "docx" "odt" "Word Document" || ((total_errors+=$?))
    convert_files "doc" "odt" "Word 97-2003 Document" || ((total_errors+=$?))
    convert_files "dotx" "ott" "Word Template" || ((total_errors+=$?))
    echo ""
    
    # === HOJAS DE CÁLCULO ===
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    log_info "║  HOJAS DE CÁLCULO                                             ║"
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    convert_files "xlsx" "ods" "Excel Workbook" || ((total_errors+=$?))
    convert_files "xls" "ods" "Excel 97-2003 Workbook" || ((total_errors+=$?))
    convert_files "xlsm" "ods" "Excel Macro-Enabled Workbook" || ((total_errors+=$?))
    convert_files "xltx" "ots" "Excel Template" || ((total_errors+=$?))
    echo ""
    
    # === PRESENTACIONES ===
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    log_info "║  PRESENTACIONES                                               ║"
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    convert_files "pptx" "odp" "PowerPoint Presentation" || ((total_errors+=$?))
    convert_files "ppt" "odp" "PowerPoint 97-2003 Presentation" || ((total_errors+=$?))
    convert_files "pptm" "odp" "PowerPoint Macro-Enabled Presentation" || ((total_errors+=$?))
    convert_files "ppsx" "odp" "PowerPoint Show" || ((total_errors+=$?))
    convert_files "pps" "odp" "PowerPoint 97-2003 Show" || ((total_errors+=$?))
    convert_files "potx" "otp" "PowerPoint Template" || ((total_errors+=$?))
    echo ""
    
    # === ELIMINACIÓN DE ARCHIVOS ORIGINALES ===
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    log_info "║  LIMPIEZA DE ARCHIVOS ORIGINALES                              ║"
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    
    if [ $total_errors -eq 0 ]; then
        log_info "Todas las conversiones fueron exitosas. Eliminando archivos originales..."
        echo ""
    else
        log_warning "Se detectaron $total_errors errores. Eliminando solo archivos convertidos exitosamente..."
        echo ""
    fi
    
    # Documentos de texto
    delete_originals "docx"
    delete_originals "doc"
    delete_originals "dotx"
    
    # Hojas de cálculo
    delete_originals "xlsx"
    delete_originals "xls"
    delete_originals "xlsm"
    delete_originals "xltx"
    
    # Presentaciones
    delete_originals "pptx"
    delete_originals "ppt"
    delete_originals "pptm"
    delete_originals "ppsx"
    delete_originals "pps"
    delete_originals "potx"
    
    echo ""
    log_info "╔═══════════════════════════════════════════════════════════════╗"
    if [ $total_errors -eq 0 ]; then
        log_info "║  ✓ PROCESO COMPLETADO EXITOSAMENTE                           ║"
    else
        log_info "║  ⚠ PROCESO COMPLETADO CON ADVERTENCIAS                       ║"
    fi
    log_info "╚═══════════════════════════════════════════════════════════════╝"
    echo ""
    
    if [ $total_errors -gt 0 ]; then
        log_warning "Total de errores: $total_errors"
        log_warning "Los archivos con errores de conversión NO fueron eliminados"
        exit 1
    fi
}

# Mostrar ayuda si se solicita
if [[ "$1" == "-h" ]] || [[ "$1" == "--help" ]]; then
    echo "Uso: $0 [directorio]"
    echo ""
    echo "Convierte recursivamente archivos de Microsoft Office a formato OpenDocument."
    echo ""
    echo "Argumentos:"
    echo "  directorio    Directorio donde buscar archivos (por defecto: directorio actual)"
    echo ""
    echo "Opciones:"
    echo "  -h, --help    Muestra esta ayuda"
    echo ""
    echo "Ejemplos:"
    echo "  $0                    # Convierte archivos en el directorio actual"
    echo "  $0 /ruta/a/archivos   # Convierte archivos en la ruta especificada"
    echo ""
    echo "Formatos soportados:"
    echo "  Documentos:     .docx, .doc, .dotx  → .odt, .ott"
    echo "  Hojas cálculo:  .xlsx, .xls, .xlsm, .xltx → .ods, .ots"
    echo "  Presentaciones: .pptx, .ppt, .pptm, .ppsx, .pps, .potx → .odp, .otp"
    echo ""
    exit 0
fi

# Ejecutar función principal
main
