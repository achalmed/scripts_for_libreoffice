#!/usr/bin/env bash
# lib/reporter.sh — Generación de reportes HTML y CSV
# Por qué existe: tener un registro visual de qué se convirtió, cuándo
# y con qué resultado permite auditar y depurar ejecuciones anteriores.

# ---------------------------------------------------------------------------
# _generate_html_report()
# Genera un reporte HTML con tabla de resultados y estadísticas.
# Argumentos: $1=ruta del archivo HTML de salida
# ---------------------------------------------------------------------------
_generate_html_report() {
    local out_file="$1"
    local timestamp
    timestamp="$(date '+%Y-%m-%d %H:%M:%S')"

    # Contar estados
    local total=0 ok=0 skip=0 fail=0 dryrun=0

    for entry in "${LOG_RESULTS[@]}"; do
        local status
        status="$(echo -e "$entry" | cut -f3)"
        ((total++))
        case "$status" in
            ok)      ((ok++)) ;;
            skip)    ((skip++)) ;;
            fail)    ((fail++)) ;;
            dry-run) ((dryrun++)) ;;
        esac
    done

    cat > "$out_file" <<HTML
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>docflow — Reporte de sesión</title>
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 2rem; background: #f5f5f5; color: #222; }
    h1   { color: #1a5276; }
    .meta { color: #555; font-size: 0.9em; margin-bottom: 1.5rem; }
    .stats { display: flex; gap: 1rem; margin-bottom: 2rem; flex-wrap: wrap; }
    .stat-box { padding: 1rem 1.5rem; border-radius: 8px; min-width: 120px; text-align: center; }
    .stat-box .num { font-size: 2rem; font-weight: bold; }
    .stat-box .lbl { font-size: 0.8rem; text-transform: uppercase; }
    .s-total  { background: #d6eaf8; }
    .s-ok     { background: #d5f5e3; }
    .s-skip   { background: #fef9e7; }
    .s-fail   { background: #fadbd8; }
    .s-dry    { background: #e8e8e8; }
    table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
    th    { background: #1a5276; color: #fff; padding: 0.75rem 1rem; text-align: left; }
    td    { padding: 0.6rem 1rem; border-bottom: 1px solid #eee; font-size: 0.9em; }
    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #f0f8ff; }
    .badge { padding: 2px 8px; border-radius: 12px; font-size: 0.8em; font-weight: bold; }
    .ok      { background: #d5f5e3; color: #1e8449; }
    .skip    { background: #fef9e7; color: #9a7d0a; }
    .fail    { background: #fadbd8; color: #922b21; }
    .dry-run { background: #e8e8e8; color: #555; }
    .path    { font-family: monospace; font-size: 0.85em; color: #555; word-break: break-all; }
  </style>
</head>
<body>
  <h1>📄 docflow — Reporte de sesión</h1>
  <div class="meta">Generado: ${timestamp} | Versión: ${VERSION:-2.0.0}</div>

  <div class="stats">
    <div class="stat-box s-total"><div class="num">${total}</div><div class="lbl">Total</div></div>
    <div class="stat-box s-ok">  <div class="num">${ok}</div>   <div class="lbl">Exitosos</div></div>
    <div class="stat-box s-skip"><div class="num">${skip}</div> <div class="lbl">Omitidos</div></div>
    <div class="stat-box s-fail"><div class="num">${fail}</div> <div class="lbl">Fallidos</div></div>
    <div class="stat-box s-dry"> <div class="num">${dryrun}</div><div class="lbl">Dry-run</div></div>
  </div>

  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>Hora</th>
        <th>Comando</th>
        <th>Estado</th>
        <th>Archivo origen</th>
        <th>Archivo destino</th>
      </tr>
    </thead>
    <tbody>
HTML

    local i=1
    for entry in "${LOG_RESULTS[@]}"; do
        local ts mode status src dst
        ts="$(echo -e "$entry" | cut -f1)"
        mode="$(echo -e "$entry" | cut -f2)"
        status="$(echo -e "$entry" | cut -f3)"
        src="$(echo -e "$entry" | cut -f4)"
        dst="$(echo -e "$entry" | cut -f5)"

        local badge_class="$status"
        [[ "$status" == "dry-run" ]] && badge_class="dry-run"

        cat >> "$out_file" <<ROW
      <tr>
        <td>${i}</td>
        <td>${ts}</td>
        <td>${mode}</td>
        <td><span class="badge ${badge_class}">${status}</span></td>
        <td class="path">$(basename "$src")</td>
        <td class="path">$(basename "$dst")</td>
      </tr>
ROW
        ((i++))
    done

    cat >> "$out_file" <<HTML
    </tbody>
  </table>
</body>
</html>
HTML

    log_success "Reporte HTML generado: $out_file"
}

# ---------------------------------------------------------------------------
# _generate_csv_report()
# Genera un CSV plano para análisis en hojas de cálculo.
# Argumentos: $1=ruta del archivo CSV de salida
# ---------------------------------------------------------------------------
_generate_csv_report() {
    local out_file="$1"
    {
        echo "timestamp,command,status,source_file,destination_file"
        for entry in "${LOG_RESULTS[@]}"; do
            # Convertir tabs a comas y envolver en comillas dobles
            echo -e "$entry" \
                | awk -F'\t' '{printf "\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n",$1,$2,$3,$4,$5}'
        done
    } > "$out_file"
    log_success "Reporte CSV generado: $out_file"
}

# ---------------------------------------------------------------------------
# _print_summary_table()
# Imprime una tabla de resumen en la terminal al final de cada sesión.
# ---------------------------------------------------------------------------
_print_summary_table() {
    [[ ${#LOG_RESULTS[@]} -eq 0 ]] && return 0

    echo ""
    log_header "RESUMEN DE CONVERSIONES"
    printf "  %-50s %-10s %-8s\n" "Archivo" "Comando" "Estado"
    printf "  %-50s %-10s %-8s\n" "$(printf '%0.s─' {1..50})" "$(printf '%0.s─' {1..10})" "$(printf '%0.s─' {1..8})"

    for entry in "${LOG_RESULTS[@]}"; do
        local mode status src
        mode="$(echo -e "$entry" | cut -f2)"
        status="$(echo -e "$entry" | cut -f3)"
        src="$(echo -e "$entry" | cut -f4)"
        local color=""
        case "$status" in
            ok)      color="$C_GREEN" ;;
            fail)    color="$C_RED" ;;
            skip)    color="$C_YELLOW" ;;
            dry-run) color="$C_BLUE" ;;
        esac
        printf "  %-50s %-10s ${color}%-8s${C_RESET}\n" \
            "$(basename "$src" | cut -c1-50)" \
            "$mode" \
            "$status"
    done
    echo ""
}

# ---------------------------------------------------------------------------
# run_report()
# Punto de entrada público del comando report.
# Genera el reporte del formato solicitado o imprime tabla si no hay resultados acumulados.
# ---------------------------------------------------------------------------
run_report() {
    log_header "GENERACIÓN DE REPORTE"

    local fmt="${REPORT_FORMAT:-html}"
    local out="${REPORT_OUT:-}"

    if [[ ${#LOG_RESULTS[@]} -eq 0 ]]; then
        log_warn "No hay resultados en la sesión actual para reportar."
        log_warn "Ejecuta primero un comando de conversión (to-odf, to-pdf, to-md)."
        exit 0
    fi

    if [[ -z "$out" ]]; then
        local ts
        ts="$(date '+%Y%m%d_%H%M%S')"
        out="${LOGS_DIR}/reporte_${ts}.${fmt}"
    fi

    validate_writable_dir "$(dirname "$out")" || exit 4

    case "$fmt" in
        html) _generate_html_report "$out" ;;
        csv)  _generate_csv_report  "$out" ;;
        *)
            log_error "Formato de reporte desconocido: '$fmt' (html | csv)"
            exit 2
            ;;
    esac

    _print_summary_table
}
