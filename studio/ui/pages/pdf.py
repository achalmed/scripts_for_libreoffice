"""Dominio: Gestión de PDF (backend: pdf-suite; PDF/A vía docflow).

Aquí vive TODA la manipulación PDF de la app — las operaciones equivalentes
que docflow y pdf-suite duplicaban están unificadas en este único módulo.

Las operaciones se declaran como especificaciones (categoría, campos,
constructor de comando); el formulario se genera dinámicamente. Añadir una
operación nueva = añadir una entrada a OPERACIONES.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDoubleSpinBox, QFormLayout, QGroupBox,
    QHBoxLayout, QLineEdit, QPushButton, QSpinBox, QVBoxLayout, QWidget,
)

from studio.services import docflow, pdfsuite
from studio.ui.pages.base import BasePage, ListaRutas, SelectorRuta

FILTRO_PDF = "PDF (*.pdf);;Todos los archivos (*)"

# ---------------------------------------------------------------------------
# Especificación de operaciones
# Campos: (clave, etiqueta, tipo, extra)
#   tipos: files | file | dir | text | password | int | float | combo | check
# El constructor recibe (valores, kw_globales) y devuelve la línea de comando.
# ---------------------------------------------------------------------------
OPERACIONES = {
    # --- Organizar ----------------------------------------------------------
    "unir": ("Organizar", "Unir PDFs", [
        ("archivos", "PDFs a unir (en orden)", "files", None),
    ], lambda v, kw: pdfsuite.unir(v["archivos"], **kw)),

    "dividir": ("Organizar", "Dividir en partes", [
        ("archivo", "PDF", "file", None),
        ("paginas", "Páginas por parte (0 = de a una)", "int", (0, 9999, 0)),
    ], lambda v, kw: pdfsuite.dividir(v["archivo"], v["paginas"], **kw)),

    "extraer": ("Organizar", "Extraer páginas", [
        ("archivo", "PDF", "file", None),
        ("paginas", "Páginas (ej. 1-5 · 1,3,10-12 · 10-z)", "text", "1-5"),
    ], lambda v, kw: pdfsuite.extraer(v["archivo"], v["paginas"], **kw)),

    "rotar": ("Organizar", "Rotar páginas", [
        ("archivo", "PDF", "file", None),
        ("angulo", "Ángulo", "combo", ("90", "180", "270", "-90")),
        ("paginas", "Páginas (vacío = todas)", "text", ""),
    ], lambda v, kw: pdfsuite.rotar(v["archivo"], int(v["angulo"]), v["paginas"], **kw)),

    "reordenar": ("Organizar", "Reordenar páginas", [
        ("archivo", "PDF", "file", None),
        ("orden", "Orden (ej. 3,1,2 · vacío = invertir)", "text", ""),
    ], lambda v, kw: pdfsuite.reordenar(v["archivo"], v["orden"], **kw)),

    "eliminar": ("Organizar", "Eliminar páginas", [
        ("archivo", "PDF", "file", None),
        ("paginas", "Páginas a eliminar (ej. 3,5,10-12)", "text", ""),
    ], lambda v, kw: pdfsuite.eliminar_paginas(v["archivo"], v["paginas"], **kw)),

    # --- Optimizar ------------------------------------------------------------
    "comprimir": ("Optimizar", "Comprimir", [
        ("rutas", "PDFs o carpetas", "files", None),
        ("metodo", "Método", "combo", ("ebook", "screen", "printer", "prepress", "ocr")),
        ("umbral", "Umbral mínimo de reducción % (0 = siempre)", "int", (0, 95, 0)),
        ("recursivo", "Recursivo en carpetas", "check", False),
    ], lambda v, kw: pdfsuite.comprimir(v["rutas"], v["metodo"], v["umbral"],
                                        recursive=v["recursivo"], **kw)),

    "test": ("Optimizar", "Comparar métodos de compresión", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.test_compresion(v["archivo"], **kw)),

    "optimizar": ("Optimizar", "Optimizar para web (linearizar)", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.optimizar(v["archivo"], **kw)),

    "reparar": ("Optimizar", "Reparar PDF dañado", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.reparar(v["archivo"], **kw)),

    "validar": ("Optimizar", "Validar estructura", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.validar(v["archivo"], **kw)),

    "pdfa": ("Optimizar", "Convertir a PDF/A-2b (archivado)", [
        ("archivos", "PDFs", "files", None),
    ], lambda v, kw: docflow.pdf_a(v["archivos"], kw.get("settings"))),  # docflow: sin equivalente en pdf-suite

    # --- Seguridad --------------------------------------------------------------
    "cifrar": ("Seguridad", "Cifrar (AES-256)", [
        ("archivo", "PDF", "file", None),
        ("owner", "Contraseña de propietario", "password", ""),
        ("user", "Contraseña de lectura (opcional)", "password", ""),
    ], lambda v, kw: pdfsuite.cifrar(v["archivo"], v["owner"], v["user"], **kw)),

    "descifrar": ("Seguridad", "Descifrar", [
        ("archivo", "PDF", "file", None),
        ("password", "Contraseña", "password", ""),
    ], lambda v, kw: pdfsuite.descifrar(v["archivo"], v["password"], **kw)),

    "marca": ("Seguridad", "Marca de agua de texto", [
        ("archivo", "PDF", "file", None),
        ("texto", "Texto", "text", "BORRADOR"),
        ("opacidad", "Opacidad (0 = defecto)", "float", (0.0, 1.0, 0.0)),
        ("angulo", "Ángulo (0 = defecto)", "int", (-360, 360, 0)),
    ], lambda v, kw: pdfsuite.marca_agua(v["archivo"], v["texto"],
                                         v["opacidad"], v["angulo"], **kw)),

    "numerar": ("Seguridad", "Añadir números de página", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.numerar_paginas(v["archivo"], **kw)),

    # --- Contenido -----------------------------------------------------------------
    "convertir": ("Contenido", "Convertir (PDF↔imagen/texto/HTML)", [
        ("rutas", "PDFs (o imágenes si el destino es pdf)", "files", None),
        ("destino", "Destino", "combo", ("png", "jpg", "svg", "txt", "html", "pdf")),
        ("dpi", "DPI (0 = defecto)", "int", (0, 1200, 0)),
    ], lambda v, kw: pdfsuite.convertir(v["rutas"], v["destino"], v["dpi"], **kw)),

    "extraer_img": ("Contenido", "Extraer imágenes embebidas", [
        ("archivo", "PDF", "file", None),
    ], lambda v, kw: pdfsuite.extraer_imagenes(v["archivo"], **kw)),

    "ocr": ("Contenido", "OCR (capa de texto buscable)", [
        ("rutas", "PDFs o carpetas", "files", None),
        ("idioma", "Idioma(s) Tesseract", "text", "spa+eng"),
        ("recursivo", "Recursivo en carpetas", "check", False),
    ], lambda v, kw: pdfsuite.ocr(v["rutas"], v["idioma"],
                                  recursive=v["recursivo"], **kw)),

    "ocr_scan": ("Contenido", "Detectar PDFs que necesitan OCR", [
        ("directorio", "Carpeta a examinar", "dir", None),
    ], lambda v, kw: pdfsuite.ocr_escanear(v["directorio"], **kw)),

    # --- Información -------------------------------------------------------------------
    "info": ("Información", "Información completa", [
        ("archivos", "PDFs", "files", None),
    ], lambda v, kw: pdfsuite.info(v["archivos"], **kw)),
}

CAMPOS_MULTIRUTA = {"files"}


class PdfPage(BasePage):
    titulo = "Gestión de PDF"
    descripcion = ("Organizar, optimizar, proteger, convertir e inspeccionar PDFs. "
                   "Equivalente de escritorio a iLovePDF, con los motores locales "
                   "qpdf, Ghostscript, poppler y ocrmypdf.")

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(task_manager, settings, parent)
        lay = self.layout_contenido()

        # Selector de operación agrupado por categoría
        self.selector = QComboBox()
        self._claves: list[str] = []
        categoria_previa = None
        for clave, (categoria, etiqueta, _, _) in OPERACIONES.items():
            if categoria != categoria_previa:
                self.selector.insertSeparator(self.selector.count()) if categoria_previa else None
                categoria_previa = categoria
            self.selector.addItem(f"{categoria} · {etiqueta}")
            self._claves.append(clave)
        self.selector.currentIndexChanged.connect(self._reconstruir_formulario)

        form_sel = QFormLayout()
        form_sel.addRow("Operación:", self.selector)
        lay.addLayout(form_sel)

        # Contenedor del formulario dinámico
        self._caja_params = QGroupBox("Parámetros")
        self._caja_params.setLayout(QVBoxLayout())
        lay.addWidget(self._caja_params)

        # Opciones globales de pdf-suite
        caja_global = QGroupBox("Opciones globales")
        fila = QHBoxLayout(caja_global)
        self.dry_run = QCheckBox("Simular (dry-run)")
        self.verbose = QCheckBox("Detallado")
        self.force = QCheckBox("Sobreescribir salidas")
        fila.addWidget(self.dry_run)
        fila.addWidget(self.verbose)
        fila.addWidget(self.force)
        fila.addStretch(1)
        lay.addWidget(caja_global)

        self.salida = SelectorRuta("pdf-out", settings, guardar=True, filtro=FILTRO_PDF)
        self.salida.campo.setPlaceholderText("(defecto: junto al original, con sufijo)")
        form_out = QFormLayout()
        form_out.addRow("Archivo de salida:", self.salida)
        lay.addLayout(form_out)

        botones = QHBoxLayout()
        botones.addStretch(1)
        ejecutar = QPushButton("Ejecutar")
        ejecutar.setDefault(True)
        ejecutar.clicked.connect(self._lanzar)
        botones.addWidget(ejecutar)
        lay.addLayout(botones)
        lay.addStretch(1)

        self._widgets: dict[str, QWidget] = {}
        self._reconstruir_formulario()

    # -- formulario dinámico ---------------------------------------------------
    def _clave_actual(self) -> str:
        # Los separadores no tienen clave: mapear índice visible → clave real
        indice = self.selector.currentIndex()
        texto = self.selector.currentText()
        for clave, (cat, etq, _, _) in OPERACIONES.items():
            if texto == f"{cat} · {etq}":
                return clave
        return self._claves[min(indice, len(self._claves) - 1)]

    def _reconstruir_formulario(self) -> None:
        contenedor = self._caja_params.layout()
        while contenedor.count():
            item = contenedor.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._widgets.clear()

        _, _, campos, _ = OPERACIONES[self._clave_actual()]
        form = QFormLayout()
        for clave, etiqueta, tipo, extra in campos:
            widget = self._crear_campo(tipo, extra)
            self._widgets[clave] = widget
            if tipo == "check":
                form.addRow("", widget)
                widget.setText(etiqueta)
            else:
                form.addRow(etiqueta + ":", widget)
        panel = QWidget()
        panel.setLayout(form)
        contenedor.addWidget(panel)

    def _crear_campo(self, tipo: str, extra) -> QWidget:
        if tipo == "files":
            return ListaRutas("pdf", self.settings, filtro=FILTRO_PDF)
        if tipo == "file":
            return SelectorRuta("pdf", self.settings, filtro=FILTRO_PDF)
        if tipo == "dir":
            return SelectorRuta("pdf", self.settings, carpeta=True)
        if tipo in ("text", "password"):
            campo = QLineEdit(str(extra or ""))
            if tipo == "password":
                campo.setEchoMode(QLineEdit.Password)
            return campo
        if tipo == "int":
            minimo, maximo, defecto = extra
            campo = QSpinBox()
            campo.setRange(minimo, maximo)
            campo.setValue(defecto)
            return campo
        if tipo == "float":
            minimo, maximo, defecto = extra
            campo = QDoubleSpinBox()
            campo.setRange(minimo, maximo)
            campo.setSingleStep(0.05)
            campo.setValue(defecto)
            return campo
        if tipo == "combo":
            campo = QComboBox()
            campo.addItems(list(extra))
            return campo
        if tipo == "check":
            return QCheckBox()
        raise ValueError(f"Tipo de campo desconocido: {tipo}")

    def _leer_valores(self) -> dict:
        valores = {}
        for clave, widget in self._widgets.items():
            if isinstance(widget, ListaRutas):
                valores[clave] = widget.rutas()
            elif isinstance(widget, SelectorRuta):
                valores[clave] = widget.ruta()
            elif isinstance(widget, QLineEdit):
                valores[clave] = widget.text().strip()
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                valores[clave] = widget.value()
            elif isinstance(widget, QComboBox):
                valores[clave] = widget.currentText()
            elif isinstance(widget, QCheckBox):
                valores[clave] = widget.isChecked()
        return valores

    # -- integración explorador / lanzamiento ------------------------------------
    def recibir_rutas(self, rutas: list[str]) -> None:
        for widget in self._widgets.values():
            if isinstance(widget, ListaRutas):
                widget.agregar(rutas)
                return
            if isinstance(widget, SelectorRuta) and rutas:
                widget.set_ruta(rutas[0])
                return

    def _lanzar(self) -> None:
        clave = self._clave_actual()
        _, etiqueta, campos, constructor = OPERACIONES[clave]
        valores = self._leer_valores()

        for c, etq, tipo, _ in campos:
            if tipo in CAMPOS_MULTIRUTA and not valores[c]:
                self.aviso(f"Falta: {etq}.")
                return
            if tipo in ("file", "dir") and not valores[c]:
                self.aviso(f"Falta: {etq}.")
                return
        if clave == "cifrar" and not valores["owner"]:
            self.aviso("La contraseña de propietario es obligatoria.")
            return

        kw = dict(settings=self.settings,
                  dry_run=self.dry_run.isChecked(),
                  verbose=self.verbose.isChecked(),
                  force=self.force.isChecked(),
                  output=self.salida.ruta())
        self.ejecutar(f"PDF · {etiqueta}", constructor(valores, kw))
