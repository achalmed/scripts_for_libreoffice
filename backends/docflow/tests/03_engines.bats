#!/usr/bin/env bats
# tests/03_engines.bats — PDF/ODF/Office engines, operaciones PDF y parsers.

load test_helper

setup() {
    setup_sandbox
    have pandoc || skip "pandoc no disponible"
}
teardown() { teardown_sandbox; }

@test "pptx → md con parser propio (títulos y viñetas)" {
    cat > slides.md <<'EOF'
# Diapositiva uno

- punto uno

# Diapositiva dos

- punto dos
EOF
    pandoc slides.md -o presentacion.pptx
    run "$DOCFLOW_BIN" to-md presentacion.pptx
    [ "$status" -eq 0 ]
    grep -q '# Diapositiva uno' presentacion.md
    grep -q -- '- punto dos' presentacion.md
}

@test "xlsx parser exporta hojas como tablas" {
    have soffice || skip "libreoffice no disponible"
    printf 'col1,col2\nuno,1\ndos,2\n' > tabla.csv
    soffice --headless --convert-to xlsx --outdir . tabla.csv >/dev/null 2>&1
    [ -f tabla.xlsx ] || skip "soffice no generó xlsx"
    run "$DOCFLOW_BIN" to-md tabla.xlsx
    [ "$status" -eq 0 ]
    grep -q '| col1 | col2 |' tabla.md
}

@test "to-pdf genera un PDF válido" {
    have soffice || skip "libreoffice no disponible"
    make_docx informe.docx
    run "$DOCFLOW_BIN" to-pdf informe.docx
    [ "$status" -eq 0 ]
    [ -s informe.pdf ]
    head -c 5 informe.pdf | grep -q '%PDF-'
}

@test "to-odf convierte docx a odt" {
    have soffice || skip "libreoffice no disponible"
    make_docx informe.docx
    run "$DOCFLOW_BIN" to-odf informe.docx
    [ "$status" -eq 0 ]
    [ -s informe.odt ]
}

@test "to-office convierte odt a docx (ida y vuelta)" {
    have soffice || skip "libreoffice no disponible"
    make_docx origen.docx
    "$DOCFLOW_BIN" to-odf origen.docx
    rm origen.docx
    run "$DOCFLOW_BIN" to-office origen.odt
    [ "$status" -eq 0 ]
    [ -s origen.docx ]
}

@test "pdf merge une dos PDFs" {
    have soffice || skip "libreoffice no disponible"
    have qpdf || skip "qpdf no disponible"
    make_docx a.docx && make_docx b.docx
    "$DOCFLOW_BIN" to-pdf a.docx b.docx
    run "$DOCFLOW_BIN" pdf merge unido.pdf a.pdf b.pdf
    [ "$status" -eq 0 ]
    [ -s unido.pdf ]
}

@test "pdf encrypt protege y docflow lo detecta como protegido" {
    have soffice || skip "libreoffice no disponible"
    have qpdf || skip "qpdf no disponible"
    make_docx a.docx
    "$DOCFLOW_BIN" to-pdf a.docx
    "$DOCFLOW_BIN" pdf encrypt a.pdf --password clave -o secreto.pdf
    run "$DOCFLOW_BIN" to-md secreto.pdf --json
    [[ "$output" == *'"protected": 1'* ]]
}

@test "--backup mueve solo originales convertidos con éxito" {
    make_docx bueno.docx
    run "$DOCFLOW_BIN" to-md bueno.docx --backup
    [ "$status" -eq 0 ]
    [ -f _backup/bueno.docx ]
    [ ! -f bueno.docx ]
    [ -f bueno.md ]
}

@test "documento OOXML corrupto se reporta como protegido, no como éxito" {
    printf 'esto no es un zip' > roto.docx
    run "$DOCFLOW_BIN" to-md roto.docx --json
    [ "$status" -eq 6 ]
    [[ "$output" == *'"protected": 1'* ]]
}

@test "index genera índice con títulos" {
    make_docx doc1.docx
    "$DOCFLOW_BIN" to-md doc1.docx
    run "$DOCFLOW_BIN" index .
    [ "$status" -eq 0 ]
    [ -f index.md ]
    grep -q 'doc1' index.md
}

@test "report genera json de la última sesión" {
    make_docx r.docx
    "$DOCFLOW_BIN" to-md r.docx
    run "$DOCFLOW_BIN" report json
    [ "$status" -eq 0 ]
    ls "${XDG_STATE_HOME}/docflow/reports/"*.json
}
