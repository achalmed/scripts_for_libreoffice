#!/bin/bash

# ============================================================================
# PDF Page Counter - Script de Instalación
# Autor: Edison Achalma
# Universidad Nacional de San Cristóbal de Huamanga
# ============================================================================

echo ""
echo "================================================================================"
echo "  📦 INSTALADOR DE PDF PAGE COUNTER"
echo "================================================================================"
echo "  👤 Autor: Edison Achalma"
echo "  🏛️  Universidad Nacional de San Cristóbal de Huamanga"
echo "================================================================================"
echo ""

# Colores para mensajes
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Función para imprimir mensajes con color
print_success() {
    echo -e "${GREEN}✓${NC} $1"
}

print_error() {
    echo -e "${RED}✗${NC} $1"
}

print_info() {
    echo -e "${BLUE}ℹ${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

# ============================================================================
# PASO 1: Verificar Python
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔍 PASO 1: Verificando Python"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

if ! command -v python3 &> /dev/null; then
    print_error "Python 3 no está instalado."
    echo ""
    print_info "Por favor, instala Python 3 primero:"
    echo "    - Ubuntu/Debian: sudo apt install python3"
    echo "    - Arch Linux: sudo pacman -S python"
    echo "    - Fedora: sudo dnf install python3"
    echo ""
    exit 1
fi

PYTHON_VERSION=$(python3 --version)
print_success "$PYTHON_VERSION encontrado"
echo ""

# ============================================================================
# PASO 2: Verificar/Instalar Conda
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🐍 PASO 2: Verificando Conda"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

if command -v conda &> /dev/null; then
    CONDA_VERSION=$(conda --version)
    print_success "$CONDA_VERSION encontrado"
    USE_CONDA=true
else
    print_warning "Conda no está instalado."
    print_info "Se usará pip para la instalación."
    USE_CONDA=false
fi
echo ""

# ============================================================================
# PASO 3: Crear/Activar Entorno
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🌐 PASO 3: Configurando Entorno"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

ENV_NAME="pdf_counter"

if [ "$USE_CONDA" = true ]; then
    print_info "Verificando entorno conda '$ENV_NAME'..."
    
    # Verificar si el entorno ya existe
    if conda env list | grep -q "^$ENV_NAME "; then
        print_success "Entorno '$ENV_NAME' ya existe"
        read -p "¿Deseas reinstalar las dependencias? (s/n): " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Ss]$ ]]; then
            print_info "Reinstalando dependencias..."
        else
            print_info "Saltando instalación de dependencias"
            echo ""
            echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            echo "✨ Para activar el entorno, ejecuta:"
            echo "   conda activate $ENV_NAME"
            echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            echo ""
            exit 0
        fi
    else
        print_info "Creando entorno conda '$ENV_NAME'..."
        conda create -n $ENV_NAME python=3.9 -y
        if [ $? -eq 0 ]; then
            print_success "Entorno '$ENV_NAME' creado exitosamente"
        else
            print_error "Error al crear el entorno conda"
            exit 1
        fi
    fi
    
    # Activar entorno (para la sesión de instalación)
    eval "$(conda shell.bash hook)"
    conda activate $ENV_NAME
    print_success "Entorno activado"
else
    print_info "Usando entorno Python del sistema"
fi
echo ""

# ============================================================================
# PASO 4: Instalar Dependencias
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📚 PASO 4: Instalando Dependencias"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

if [ "$USE_CONDA" = true ]; then
    print_info "Instalando PyPDF2 con conda..."
    conda install -c conda-forge pypdf2 -y
    
    print_info "Instalando openpyxl con conda..."
    conda install -c conda-forge openpyxl -y
else
    print_info "Instalando PyPDF2 con pip..."
    pip3 install PyPDF2 --break-system-packages
    
    print_info "Instalando openpyxl con pip..."
    pip3 install openpyxl --break-system-packages
fi

if [ $? -eq 0 ]; then
    print_success "Dependencias instaladas correctamente"
else
    print_error "Hubo problemas instalando dependencias"
    echo ""
    print_info "Intenta manualmente:"
    if [ "$USE_CONDA" = true ]; then
        echo "    conda install -c conda-forge pypdf2 openpyxl"
    else
        echo "    pip3 install PyPDF2 openpyxl --break-system-packages"
    fi
    exit 1
fi
echo ""

# ============================================================================
# PASO 5: Configurar Permisos
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔐 PASO 5: Configurando Permisos"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

chmod +x main.py
if [ $? -eq 0 ]; then
    print_success "Permisos de ejecución configurados"
else
    print_warning "No se pudieron configurar los permisos"
fi
echo ""

# ============================================================================
# PASO 6: Crear Directorio para Excel
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📁 PASO 6: Creando Directorio de Datos"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

mkdir -p excel_databases
if [ $? -eq 0 ]; then
    print_success "Directorio 'excel_databases' creado"
else
    print_warning "No se pudo crear el directorio"
fi
echo ""

# ============================================================================
# PASO 7: Verificar Instalación
# ============================================================================
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🧪 PASO 7: Verificando Instalación"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

python3 -c "import PyPDF2; import openpyxl; print('✅ Todas las bibliotecas están funcionando correctamente')" 2>/dev/null
if [ $? -eq 0 ]; then
    print_success "Verificación exitosa"
else
    print_error "Error en la verificación"
    exit 1
fi
echo ""

# ============================================================================
# FINALIZACIÓN
# ============================================================================
echo ""
echo "================================================================================"
echo "  ✨ INSTALACIÓN COMPLETADA EXITOSAMENTE"
echo "================================================================================"
echo ""

if [ "$USE_CONDA" = true ]; then
    echo "📖 INSTRUCCIONES DE USO CON CONDA:"
    echo ""
    echo "   1. Activa el entorno conda:"
    echo "      ${GREEN}conda activate $ENV_NAME${NC}"
    echo ""
    echo "   2. Ejecuta el script:"
    echo "      ${GREEN}python3 main.py${NC}"
    echo ""
    echo "   3. Para desactivar el entorno:"
    echo "      ${GREEN}conda deactivate${NC}"
    echo ""
else
    echo "📖 INSTRUCCIONES DE USO:"
    echo ""
    echo "   Ejecuta el script:"
    echo "   ${GREEN}python3 main.py${NC}"
    echo ""
fi

echo "📚 EJEMPLOS:"
echo "   • Ver blogs disponibles:"
echo "     ${BLUE}python3 main.py --listar${NC}"
echo ""
echo "   • Contar todos los blogs:"
echo "     ${BLUE}python3 main.py${NC}"
echo ""
echo "   • Contar blogs específicos:"
echo "     ${BLUE}python3 main.py -b actus-mercator aequilibria${NC}"
echo ""
echo "   • Ver ayuda completa:"
echo "     ${BLUE}python3 main.py --help${NC}"
echo ""
echo "📄 Para más información, consulta README.md"
echo ""
echo "================================================================================"
echo ""

# Crear alias sugerido
echo "💡 TIP: Puedes crear un alias para facilitar el uso:"
echo ""
if [ "$USE_CONDA" = true ]; then
    echo "   Agrega esto a tu ~/.bashrc o ~/.zshrc:"
    echo "   ${YELLOW}alias count-pdfs='conda activate $ENV_NAME && python3 $(pwd)/main.py'${NC}"
else
    echo "   Agrega esto a tu ~/.bashrc o ~/.zshrc:"
    echo "   ${YELLOW}alias count-pdfs='python3 $(pwd)/main.py'${NC}"
fi
echo ""