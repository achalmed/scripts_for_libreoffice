"""Dominio: Conversión de documentos (backend: docflow).

Un solo lugar para to-md / to-pdf / to-odf / to-office, con las opciones
relevantes según el destino. «Simular» ejecuta el mismo lote con --dry-run.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFormLayout, QGroupBox, QHBoxLayout, QLineEdit,
    QPushButton, QSpinBox,
)

from studio.services import docflow
from studio.ui.pages.base import BasePage, ListaRutas, SelectorRuta

DESTINOS = (
    ("Markdown (to-md)", "to-md"),
    ("PDF (to-pdf)", "to-pdf"),
    ("OpenDocument (to-odf)", "to-odf"),
    ("MS Office (to-office)", "to-office"),
)
FLAVORS = ("gfm", "obsidian", "logseq", "mkdocs", "hugo", "quarto")


class ConversionPage(BasePage):
    titulo = "Conversión de documentos"
    descripcion = ("Office, OpenDocument, PDF, EPUB, HTML y LaTeX hacia Markdown, "
                   "PDF u Office↔ODF. Los originales nunca se tocan salvo que lo pidas.")

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(task_manager, settings, parent)
        lay = self.layout_contenido()

        self.entradas = ListaRutas("conversion", settings)
        caja_entrada = QGroupBox("Entrada (archivos o carpetas)")
        caja_entrada.setLayout(QHBoxLayout())
        caja_entrada.layout().addWidget(self.entradas)
        lay.addWidget(caja_entrada)

        # -- destino y opciones generales -------------------------------------
        form = QFormLayout()
        self.destino = QComboBox()
        for etiqueta, _ in DESTINOS:
            self.destino.addItem(etiqueta)
        self.destino.currentIndexChanged.connect(self._actualizar_grupos)
        form.addRow("Convertir a:", self.destino)

        self.salida = SelectorRuta("conversion-out", settings, carpeta=True)
        self.salida.campo.setPlaceholderText("(defecto: junto a cada original)")
        form.addRow("Carpeta de salida:", self.salida)

        self.formatos = QLineEdit()
        self.formatos.setPlaceholderText("ej.: docx,pptx (vacío = todos los soportados)")
        form.addRow("Solo extensiones:", self.formatos)

        self.excluir = QLineEdit()
        self.excluir.setPlaceholderText("ej.: node_modules, *.tmp (separados por coma)")
        form.addRow("Excluir:", self.excluir)

        self.jobs = QSpinBox()
        self.jobs.setRange(0, 64)
        self.jobs.setSpecialValueText("auto (nproc)")
        form.addRow("Trabajos en paralelo:", self.jobs)

        opciones = QHBoxLayout()
        self.overwrite = QCheckBox("Sobreescribir salidas")
        self.resume = QCheckBox("Reanudar lote")
        self.no_cache = QCheckBox("Sin caché")
        self.backup = QCheckBox("Backup de originales exitosos")
        for w in (self.overwrite, self.resume, self.no_cache, self.backup):
            opciones.addWidget(w)
        opciones.addStretch(1)

        self.reporte = QComboBox()
        self.reporte.addItems(["(sin reporte)", "html", "csv", "json", "md", "html,json"])
        form.addRow("Reporte al terminar:", self.reporte)

        caja_general = QGroupBox("Opciones del lote")
        cuerpo = QFormLayout(caja_general)
        cuerpo.addRow(form)
        cuerpo.addRow(opciones)
        lay.addWidget(caja_general)

        # -- opciones Markdown --------------------------------------------------
        self.grupo_md = QGroupBox("Markdown")
        form_md = QFormLayout(self.grupo_md)
        self.flavor = QComboBox()
        self.flavor.addItems(FLAVORS)
        form_md.addRow("Sabor (flavor):", self.flavor)
        self.sin_metadata = QCheckBox("Sin frontmatter YAML")
        self.sin_notas = QCheckBox("Sin notas del presentador")
        fila_md = QHBoxLayout()
        fila_md.addWidget(self.sin_metadata)
        fila_md.addWidget(self.sin_notas)
        fila_md.addStretch(1)
        form_md.addRow(fila_md)
        self.img_format = QComboBox()
        self.img_format.addItems(["(original)", "webp", "png", "jpg"])
        form_md.addRow("Formato de imágenes:", self.img_format)
        self.img_quality = QSpinBox()
        self.img_quality.setRange(0, 100)
        self.img_quality.setSpecialValueText("defecto")
        form_md.addRow("Calidad de imagen:", self.img_quality)
        self.ocr = QCheckBox("OCR previo en PDFs escaneados")
        self.ocr_lang = QLineEdit("spa+eng")
        fila_ocr = QHBoxLayout()
        fila_ocr.addWidget(self.ocr)
        fila_ocr.addWidget(self.ocr_lang)
        form_md.addRow(fila_ocr)
        lay.addWidget(self.grupo_md)

        # -- opciones PDF --------------------------------------------------------
        self.grupo_pdf = QGroupBox("PDF")
        form_pdf = QFormLayout(self.grupo_pdf)
        self.pdfa = QCheckBox("Post-proceso PDF/A-2b")
        self.comprimir = QCheckBox("Comprimir cada PDF generado")
        fila_pdf = QHBoxLayout()
        fila_pdf.addWidget(self.pdfa)
        fila_pdf.addWidget(self.comprimir)
        fila_pdf.addStretch(1)
        form_pdf.addRow(fila_pdf)
        lay.addWidget(self.grupo_pdf)

        # -- acciones ------------------------------------------------------------
        botones = QHBoxLayout()
        botones.addStretch(1)
        simular = QPushButton("Simular (dry-run)")
        simular.clicked.connect(lambda: self._lanzar(dry_run=True))
        convertir = QPushButton("Convertir")
        convertir.setDefault(True)
        convertir.clicked.connect(lambda: self._lanzar(dry_run=False))
        botones.addWidget(simular)
        botones.addWidget(convertir)
        lay.addLayout(botones)
        lay.addStretch(1)

        self._actualizar_grupos()

    # -- lógica -------------------------------------------------------------
    def recibir_rutas(self, rutas: list[str]) -> None:
        self.entradas.agregar(rutas)

    def _comando_actual(self) -> str:
        return DESTINOS[self.destino.currentIndex()][1]

    def _actualizar_grupos(self) -> None:
        cmd = self._comando_actual()
        self.grupo_md.setVisible(cmd == "to-md")
        self.grupo_pdf.setVisible(cmd == "to-pdf")

    def _lanzar(self, dry_run: bool) -> None:
        rutas = self.entradas.rutas()
        if not rutas:
            self.aviso("Añade al menos un archivo o carpeta de entrada.")
            return
        cmd = self._comando_actual()
        op = docflow.OpcionesConversion(
            output_dir=self.salida.ruta(),
            formats=self.formatos.text().strip(),
            exclude=[e.strip() for e in self.excluir.text().split(",") if e.strip()],
            jobs=self.jobs.value(),
            overwrite=self.overwrite.isChecked(),
            dry_run=dry_run,
            resume=self.resume.isChecked(),
            no_cache=self.no_cache.isChecked(),
            backup=self.backup.isChecked(),
            report="" if self.reporte.currentIndex() == 0 else self.reporte.currentText(),
        )
        if cmd == "to-md":
            op.flavor = self.flavor.currentText()
            op.no_metadata = self.sin_metadata.isChecked()
            op.no_notes = self.sin_notas.isChecked()
            if self.img_format.currentIndex() > 0:
                op.img_format = self.img_format.currentText()
            op.img_quality = self.img_quality.value()
            op.ocr = self.ocr.isChecked()
            if op.ocr:
                op.ocr_lang = self.ocr_lang.text().strip()
        elif cmd == "to-pdf":
            op.pdfa = self.pdfa.isChecked()
            op.compress = self.comprimir.isChecked()

        etiqueta = "Simulación" if dry_run else "Conversión"
        self.ejecutar(f"{etiqueta} {cmd} ({len(rutas)} ruta/s)",
                      docflow.convertir(cmd, rutas, op, self.settings))
