#!/usr/bin/env bash
# lib/core/backup.sh — Gestión segura de archivos originales tras convertir.
# Invariante de seguridad: SOLO se tocan originales cuyo estado en la sesión
# es "ok" (conversión exitosa Y verificada). Los fallidos nunca se mueven
# ni eliminan.

# backup_handle_originals — aplica --backup o --delete según configuración
backup_handle_originals() {
    [[ "${DOCFLOW_BACKUP_ORIGINALS:-false}" == "true" ]] && { _backup_move; return 0; }
    [[ "${DOCFLOW_DELETE_ORIGINALS:-false}" == "true" ]] && { _backup_delete; return 0; }
    return 0
}

_backup_move() {
    local count
    count="$(session_inputs_with_status ok | wc -l)"
    (( count == 0 )) && return 0

    log_header "Moviendo ${count} original(es) a ${DOCFLOW_BACKUP_DIR_NAME}/"
    local file bdir
    while IFS= read -r file; do
        [[ -f "$file" ]] || continue
        bdir="$(dirname -- "$file")/${DOCFLOW_BACKUP_DIR_NAME}"
        if [[ "${DOCFLOW_DRY_RUN:-false}" == "true" ]]; then
            log_skip "[dry-run] mv $(basename -- "$file") → ${DOCFLOW_BACKUP_DIR_NAME}/"
        else
            mkdir -p -- "$bdir" && mv -- "$file" "$bdir/" \
                && log_ok "backup: $(basename -- "$file")"
        fi
    done < <(session_inputs_with_status ok)
}

_backup_delete() {
    local count
    count="$(session_inputs_with_status ok | wc -l)"
    (( count == 0 )) && return 0

    log_header "Eliminación de originales"
    if ! util_confirm "¿Eliminar ${count} archivo(s) original(es) convertidos con éxito?"; then
        log_warn "Originales conservados."
        return 0
    fi
    local file
    while IFS= read -r file; do
        [[ -f "$file" ]] || continue
        if [[ "${DOCFLOW_DRY_RUN:-false}" == "true" ]]; then
            log_skip "[dry-run] rm $(basename -- "$file")"
        else
            rm -- "$file" && log_ok "eliminado: $(basename -- "$file")"
        fi
    done < <(session_inputs_with_status ok)
}
