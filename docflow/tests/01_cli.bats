#!/usr/bin/env bats
# tests/01_cli.bats — CLI: ayuda, versión, errores de uso y códigos de salida.

load test_helper

setup()    { setup_sandbox; }
teardown() { teardown_sandbox; }

@test "version imprime nombre y versión" {
    run "$DOCFLOW_BIN" version
    [ "$status" -eq 0 ]
    [[ "$output" == docflow\ v* ]]
}

@test "sin argumentos muestra ayuda y sale con código de uso (2)" {
    run "$DOCFLOW_BIN"
    [ "$status" -eq 2 ]
}

@test "help sale con 0" {
    run "$DOCFLOW_BIN" help
    [ "$status" -eq 0 ]
}

@test "comando desconocido sale con código de uso (2)" {
    run "$DOCFLOW_BIN" comando-inventado
    [ "$status" -eq 2 ]
}

@test "flag desconocido sale con código de uso (2)" {
    run "$DOCFLOW_BIN" to-md --flag-inventado x
    [ "$status" -eq 2 ]
}

@test "to-md sin rutas sale con código de uso (2)" {
    run "$DOCFLOW_BIN" to-md
    [ "$status" -eq 2 ]
}

@test "ruta inexistente sale con EXIT_BAD_INPUT (3)" {
    run "$DOCFLOW_BIN" to-md /ruta/que/no/existe
    [ "$status" -eq 3 ]
}

@test "directorio sin archivos convertibles sale con EXIT_NO_FILES (4)" {
    mkdir -p vacio
    run "$DOCFLOW_BIN" to-md vacio/
    [ "$status" -eq 4 ]
}

@test "formats lista los formatos registrados" {
    run "$DOCFLOW_BIN" formats
    [ "$status" -eq 0 ]
    [[ "$output" == *docx* ]]
    [[ "$output" == *pptx* ]]
    [[ "$output" == *odt* ]]
}

@test "doctor se ejecuta" {
    run "$DOCFLOW_BIN" doctor
    [[ "$status" -eq 0 || "$status" -eq 5 ]]
}

@test "config init crea la plantilla TOML" {
    run "$DOCFLOW_BIN" config init
    [ "$status" -eq 0 ]
    [ -f "${XDG_CONFIG_HOME}/docflow/config.toml" ]
}

@test "config path imprime la ruta" {
    run "$DOCFLOW_BIN" config path
    [ "$status" -eq 0 ]
    [[ "$output" == *config.toml ]]
}
