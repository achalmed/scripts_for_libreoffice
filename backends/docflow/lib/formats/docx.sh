#!/usr/bin/env bash
# lib/formats/docx.sh — Word moderno (OOXML): docx y docm.
#            ext    categoría   descripción                       md      pdf      odf   ms
registry_add docx  "document"  "Word (OOXML)"                    pandoc  soffice  odt   ""
registry_add docm  "document"  "Word con macros (OOXML)"         pandoc  soffice  odt   ""
