# Autocompletado fish para docflow
complete -c docflow -f -n '__fish_use_subcommand' -a 'to-md' -d 'Convertir a Markdown'
complete -c docflow -f -n '__fish_use_subcommand' -a 'to-pdf' -d 'Convertir a PDF'
complete -c docflow -f -n '__fish_use_subcommand' -a 'to-odf' -d 'MS Office → OpenDocument'
complete -c docflow -f -n '__fish_use_subcommand' -a 'to-office' -d 'OpenDocument → MS Office'
complete -c docflow -f -n '__fish_use_subcommand' -a 'pdf' -d 'Operaciones sobre PDFs'
complete -c docflow -f -n '__fish_use_subcommand' -a 'watch' -d 'Vigilar directorio'
complete -c docflow -f -n '__fish_use_subcommand' -a 'report' -d 'Reporte de sesión'
complete -c docflow -f -n '__fish_use_subcommand' -a 'index' -d 'Índice de documentos'
complete -c docflow -f -n '__fish_use_subcommand' -a 'doctor' -d 'Diagnóstico de dependencias'
complete -c docflow -f -n '__fish_use_subcommand' -a 'formats' -d 'Formatos soportados'
complete -c docflow -f -n '__fish_use_subcommand' -a 'cache' -d 'Gestión del caché'
complete -c docflow -f -n '__fish_use_subcommand' -a 'config' -d 'Configuración'

complete -c docflow -f -n '__fish_seen_subcommand_from pdf' -a 'merge split rotate extract compress pdfa encrypt decrypt watermark ocr info'
complete -c docflow -f -n '__fish_seen_subcommand_from cache' -a 'stats clear prune'
complete -c docflow -f -n '__fish_seen_subcommand_from config' -a 'init show path'

complete -c docflow -s o -l output-dir -d 'Directorio de salida' -r
complete -c docflow -s f -l formats -d 'Extensiones (csv)' -r
complete -c docflow -s e -l exclude -d 'Excluir patrón' -r
complete -c docflow -s j -l jobs -d 'Workers en paralelo' -r
complete -c docflow -s w -l overwrite -d 'Sobreescribir salidas'
complete -c docflow -s n -l dry-run -d 'Simular sin escribir'
complete -c docflow -s v -l verbose -d 'Salida detallada'
complete -c docflow -s q -l quiet -d 'Modo silencioso'
complete -c docflow -l json -d 'Salida JSON'
complete -c docflow -l flavor -a 'gfm obsidian logseq mkdocs hugo quarto' -r
complete -c docflow -l img-format -a 'keep png jpg webp avif' -r
complete -c docflow -l pdf-engine -a 'auto soffice pandoc wkhtmltopdf weasyprint' -r
complete -c docflow -l report -a 'html csv json md' -r
complete -c docflow -l backup -d 'Mover originales a _backup/'
complete -c docflow -l delete -d 'Eliminar originales exitosos'
complete -c docflow -l force -d 'Sin confirmaciones'
complete -c docflow -l resume -d 'Reanudar lote interrumpido'
complete -c docflow -l no-cache -d 'Ignorar caché'
complete -c docflow -l ocr -d 'OCR para PDFs escaneados'
