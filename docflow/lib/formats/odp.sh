#!/usr/bin/env bash
# lib/formats/odp.sh — Presentaciones OpenDocument: odp, fodp y otp.
registry_add odp   "presentation" "Presentación OpenDocument"       via_pptx soffice "" pptx
registry_add fodp  "presentation" "Presentación OpenDocument plana" via_pptx soffice "" pptx
registry_add otp   "presentation" "Plantilla presentación ODF"      via_pptx soffice "" potx
