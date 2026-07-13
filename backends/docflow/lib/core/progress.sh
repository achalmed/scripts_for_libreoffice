#!/usr/bin/env bash
# lib/core/progress.sh — Barra de progreso con ETA.
# Solo se dibuja en TTY, sin verbose y sin quiet: en pipes/cron no ensucia.

# progress_enabled → 0 si procede dibujar la barra
progress_enabled() {
    [[ -t 2 ]] || return 1
    (( _LOG_LEVEL == 2 )) || return 1   # ni verbose ni quiet
    return 0
}

# progress_draw <actual> <total> <inicio_ms>
progress_draw() {
    local current="$1" total="$2" start_ms="$3"
    (( total == 0 )) && return 0
    local width=30 filled i bar="" eta_str=""
    filled=$(( current * width / total ))

    for (( i = 0; i < width; i++ )); do
        (( i < filled )) && bar+="█" || bar+="░"
    done

    if (( current > 0 && current < total )); then
        local elapsed=$(( $(util_epoch_ms) - start_ms ))
        local remaining=$(( elapsed * (total - current) / current / 1000 ))
        eta_str="$(printf '  ETA %02d:%02d' $((remaining/60)) $((remaining%60)))"
    fi

    printf '\r  %b%s%b  %d / %d%s ' \
        "$C_GREEN" "$bar" "$C_RESET" "$current" "$total" "$eta_str" >&2
    (( current >= total )) && printf '\n' >&2
    return 0
}
