#!/usr/bin/env bash
# lib/engines/media.sh — Media Engine: procesamiento de imágenes extraídas.
# Conversión de formato (png/jpg/webp/avif), redimensionado, compresión y
# optimización. Usa ImageMagick si está; cwebp como alternativa para WebP.
# Sin herramientas disponibles se degrada a no-op con aviso (nunca falla).

# _media_magick — nombre del binario de ImageMagick (v7: magick; v6: convert)
_media_magick() {
    if deps_has magick; then echo "magick"
    elif deps_has convert; then echo "convert"
    else echo ""; fi
}

# media_process_dir <dir_imágenes> <md_asociado>
# Aplica formato/calidad/redimensionado según configuración y actualiza los
# enlaces del Markdown si cambian las extensiones.
media_process_dir() {
    local img_dir="$1" md_file="$2"
    [[ -d "$img_dir" ]] || return 0

    local target="${DOCFLOW_IMG_FORMAT:-keep}"
    local quality="${DOCFLOW_IMG_QUALITY:-90}"
    local max_width="${DOCFLOW_IMG_MAX_WIDTH:-0}"
    local optimize="${DOCFLOW_IMG_OPTIMIZE:-false}"

    [[ "$target" == "keep" && "$max_width" == "0" && "$optimize" != "true" ]] && return 0

    local magick renamed=false img
    magick="$(_media_magick)"
    if [[ -z "$magick" ]] && ! { [[ "$target" == "webp" ]] && deps_has cwebp; }; then
        log_warn "ImageMagick no disponible: imágenes sin procesar (instala: $(deps_install_hint magick))"
        return 0
    fi

    while IFS= read -r -d '' img; do
        local ext="${img##*.}" new_img="$img"
        ext="${ext,,}"

        # Conversión de formato
        if [[ "$target" != "keep" && "$ext" != "$target" && "$ext" != "svg" ]]; then
            new_img="${img%.*}.${target}"
            if [[ "$target" == "webp" && -z "$magick" ]]; then
                cwebp -quiet -q "$quality" "$img" -o "$new_img" 2>/dev/null || { new_img="$img"; continue; }
            else
                "$magick" "$img" -quality "$quality" "$new_img" 2>/dev/null || { new_img="$img"; continue; }
            fi
            [[ "$new_img" != "$img" && -s "$new_img" ]] && { rm -f -- "$img"; renamed=true; }
        fi

        # Redimensionado (solo reduce, nunca amplía: flag '>')
        if [[ "$max_width" != "0" && -n "$magick" && -f "$new_img" ]]; then
            "$magick" "$new_img" -resize "${max_width}x>" "$new_img" 2>/dev/null
        fi

        # Optimización sin cambiar formato
        if [[ "$optimize" == "true" && -n "$magick" && -f "$new_img" ]]; then
            "$magick" "$new_img" -strip -quality "$quality" "$new_img" 2>/dev/null
        fi
    done < <(find "$img_dir" -type f \
        \( -iname '*.png' -o -iname '*.jpg' -o -iname '*.jpeg' \
           -o -iname '*.gif' -o -iname '*.webp' -o -iname '*.avif' \
           -o -iname '*.bmp' -o -iname '*.tiff' -o -iname '*.emf' -o -iname '*.wmf' \) -print0)

    # Si cambió alguna extensión, reapuntar los enlaces del Markdown
    if [[ "$renamed" == "true" && -f "$md_file" ]]; then
        python3 "${DOCFLOW_ROOT}/lib/helpers/mdtools.py" retarget-media "$md_file" \
            --media-dir "$img_dir" || true
    fi
    return 0
}
