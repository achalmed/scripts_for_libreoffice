#!/usr/bin/env bash
# lib/formats/ppt.sh — PowerPoint legado (binario 97-2003).
# via_pptx: LibreOffice lo moderniza a pptx y luego actúa el parser propio.
registry_add ppt   "presentation" "PowerPoint 97-2003"           via_pptx soffice odp ""
registry_add pps   "presentation" "Presentación 97-2003 (show)"  via_pptx soffice odp ""
registry_add pot   "presentation" "Plantilla PowerPoint 97-2003" via_pptx soffice otp ""
