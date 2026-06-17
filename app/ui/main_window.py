from __future__ import annotations

import csv
import os
import shutil
import sqlite3
import subprocess
import sys
import uuid
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from pydantic import ValidationError
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.core.settings import (
    CUSTOMER_CATALOG_PATH,
    EXCEL_TEMPLATES_DIR,
    INPUT_DIR,
    LOGO_DIR,
    MASTER_DATA_EXPORT_PATH,
    OUTPUT_DIR,
    PDF_TEMPLATES_DIR,
    SIGNATURE_DIR,
    WORD_TEMPLATES_DIR,
    ensure_directories,
    logo_path,
    signature_path,
)
from app.database.database_service import DatabaseService
from app.models.schemas import MasterData, TemplateRecord
from app.services.ai_mapping_service import AiMappingService
from app.services.customer_catalog_service import (
    CUSTOMER_KEYS,
    CustomerCatalogService,
)
from app.services.form_import_service import FormImportService
from app.services.form_template_service import FormTemplateService
from app.services.mapping_service import MappingService
from app.services.pdf_field_service import PdfFieldService
from app.services.pdf_fill_service import PdfFillService
from app.ui.import_preview_dialog import ImportPreviewDialog
from app.ui.pdf_learning_dialog import PdfLearningDialog
from app.ui.theme import APP_STYLESHEET


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        ensure_directories()
        self.db = DatabaseService()
        self.db.initialize()
        self.pdf_fields = PdfFieldService()
        self.mapping_service = MappingService()
        self.pdf_fill = PdfFillService()
        self.ai_mapping = AiMappingService()
        self.customer_catalog = CustomerCatalogService()
        self.form_import = FormImportService()
        self.form_templates = FormTemplateService()

        self.current_pdf: Path | None = None
        self.current_template_id: int | None = None
        self.current_reference_path: Path | None = None
        self.current_fields: list[str] = []
        self.current_mapping_path: Path | None = None
        self.current_mapping_payload: dict[str, object] | None = None
        self.selected_master_id: int | None = None
        self.selected_certificate_customer: dict[str, object] | None = None
        self._load_customer_catalog_if_available()

        self.setWindowTitle("Madecentro | Gestión inteligente de formularios")
        self.setMinimumSize(1100, 720)
        self.resize(1360, 860)
        self.setStyleSheet(APP_STYLESHEET)

        central = QWidget()
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)
        central_layout.addWidget(self._build_brand_header())

        self.tabs = QTabWidget()
        central_layout.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

        self._build_master_tab()
        self._build_templates_tab()
        self._build_mapping_tab()
        self._build_fill_tab()
        self._build_certificate_tab()
        self._build_instructions_tab()
        self.refresh_all()
        self._sync_master_data_file()

    def _build_brand_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("brandHeader")
        header.setFixedHeight(88)
        layout = QHBoxLayout(header)
        layout.setContentsMargins(28, 15, 28, 15)
        layout.setSpacing(16)

        self.logo_label = QLabel()
        self.logo_label.setFixedSize(185, 56)
        self.logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self._refresh_logo()

        divider = QFrame()
        divider.setObjectName("brandDivider")
        divider.setFixedSize(3, 42)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        title = QLabel("Gestión inteligente de formularios")
        title.setObjectName("brandTitle")
        subtitle = QLabel(
            "Automatización documental y administración de datos maestros"
        )
        subtitle.setObjectName("brandSubtitle")
        text_layout.addWidget(title)
        text_layout.addWidget(subtitle)

        layout.addWidget(self.logo_label)
        layout.addWidget(divider)
        layout.addLayout(text_layout)
        layout.addStretch()
        return header

    def _refresh_logo(self) -> None:
        path = logo_path()
        if path:
            pixmap = QPixmap(str(path))
            self.setWindowIcon(QIcon(str(path)))
            self.logo_label.setPixmap(
                pixmap.scaled(
                    self.logo_label.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation,
                )
            )
            self.logo_label.setText("")
            return
        self.logo_label.setPixmap(QPixmap())
        self.logo_label.setText("MADECENTRO")
        self.logo_label.setStyleSheet(
            "color: #F08419; font-size: 18px; font-weight: 800;"
        )

    def _page_intro(self, title: str, subtitle: str) -> QVBoxLayout:
        layout = QVBoxLayout()
        layout.setSpacing(3)
        title_label = QLabel(title)
        title_label.setObjectName("pageTitle")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("pageSubtitle")
        subtitle_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        return layout

    def _card(self) -> tuple[QFrame, QVBoxLayout]:
        frame = QFrame()
        frame.setObjectName("card")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)
        return frame, layout

    def _section_title(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("sectionTitle")
        return label

    def _prepare_table(self, table: QTableWidget) -> None:
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setShowGrid(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setHighlightSections(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)

    def _build_master_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Datos maestros",
                "Centraliza la información que se reutiliza al diligenciar cada formulario.",
            )
        )

        import_card, import_layout = self._card()
        import_layout.addWidget(self._section_title("Importar información"))
        import_actions = QHBoxLayout()
        import_btn = QPushButton("Importar CSV/XLSX maestros")
        import_customers_btn = QPushButton("Importar clientes certificados")
        import_form_btn = QPushButton("Importar formulario diligenciado")
        history_btn = QPushButton("Historial de importaciones")
        import_form_btn.setProperty("role", "primary")
        import_actions.addWidget(import_btn)
        import_actions.addWidget(import_customers_btn)
        import_actions.addWidget(import_form_btn)
        import_actions.addWidget(history_btn)
        import_actions.addStretch()
        import_layout.addLayout(import_actions)
        layout.addWidget(import_card)

        edit_card, edit_layout = self._card()
        edit_layout.addWidget(self._section_title("Crear o editar un dato"))
        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)
        save_btn = QPushButton("Guardar fila")
        edit_btn = QPushButton("Editar fila seleccionada")
        delete_btn = QPushButton("Eliminar fila seleccionada")
        save_btn.setProperty("role", "primary")
        delete_btn.setProperty("role", "danger")
        self.master_key_input = QLineEdit()
        self.master_key_input.setPlaceholderText("Ej. nit")
        self.master_value_input = QLineEdit()
        self.master_value_input.setPlaceholderText("Valor del dato maestro")
        self.master_category_input = QLineEdit("general")
        form.addWidget(QLabel("Clave"), 0, 0)
        form.addWidget(QLabel("Valor"), 0, 1)
        form.addWidget(QLabel("Categoría"), 0, 2)
        form.addWidget(self.master_key_input, 1, 0)
        form.addWidget(self.master_value_input, 1, 1)
        form.addWidget(self.master_category_input, 1, 2)
        form.setColumnStretch(1, 2)
        edit_layout.addLayout(form)
        edit_actions = QHBoxLayout()
        edit_actions.addStretch()
        edit_actions.addWidget(edit_btn)
        edit_actions.addWidget(delete_btn)
        edit_actions.addWidget(save_btn)
        edit_layout.addLayout(edit_actions)
        layout.addWidget(edit_card)

        self.master_table = QTableWidget(0, 4)
        self.master_table.setHorizontalHeaderLabels(["ID", "Clave", "Valor", "Categoria"])
        self._prepare_table(self.master_table)
        table_card, table_layout = self._card()
        table_layout.addWidget(self._section_title("Información registrada"))
        table_layout.addWidget(self.master_table)
        layout.addWidget(table_card, 1)

        import_btn.clicked.connect(self.import_master_data)
        import_customers_btn.clicked.connect(self.import_certificate_customers)
        import_form_btn.clicked.connect(self.import_completed_form)
        history_btn.clicked.connect(self.show_import_history)
        save_btn.clicked.connect(self.save_master_row)
        edit_btn.clicked.connect(self.edit_selected_master_row)
        delete_btn.clicked.connect(self.delete_selected_master_row)
        self.master_table.cellClicked.connect(self.load_master_row_from_selection)
        self.master_table.itemSelectionChanged.connect(self.load_selected_master_row)
        self.tabs.addTab(tab, "Datos maestros")

    def _build_templates_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Plantillas",
                "Registra una versión vacía y una diligenciada para aprender su estructura.",
            )
        )

        register_card, register_layout = self._card()
        register_layout.addWidget(self._section_title("Registrar formulario"))
        actions = QGridLayout()
        learn_pdf_btn = QPushButton("Aprender plantilla")
        load_btn = QPushButton("Registrar vacia + diligenciada")
        list_btn = QPushButton("Listar campos detectados")
        delete_btn = QPushButton("Eliminar plantilla seleccionada")
        learn_pdf_btn.setProperty("role", "primary")
        delete_btn.setProperty("role", "danger")
        self.template_name_input = QLineEdit()
        self.template_name_input.setPlaceholderText("Ej. Formulario de vinculación")
        actions.addWidget(QLabel("Nombre de la plantilla"), 0, 0, 1, 2)
        actions.addWidget(self.template_name_input, 1, 0)
        actions.addWidget(learn_pdf_btn, 1, 1)
        actions.addWidget(load_btn, 1, 2)
        actions.addWidget(list_btn, 1, 3)
        actions.addWidget(delete_btn, 1, 4)
        actions.setColumnStretch(0, 1)
        register_layout.addLayout(actions)
        layout.addWidget(register_card)

        self.template_table = QTableWidget(0, 5)
        self.template_table.setHorizontalHeaderLabels(
            ["ID", "Nombre", "Formato", "Plantilla vacia", "Referencia"]
        )
        self.field_table = QTableWidget(0, 4)
        self.field_table.setHorizontalHeaderLabels(
            ["Campo", "Tipo", "Ubicacion", "Valor de referencia"]
        )
        self._prepare_table(self.template_table)
        self._prepare_table(self.field_table)

        templates_card, templates_layout = self._card()
        templates_layout.addWidget(self._section_title("Plantillas registradas"))
        templates_layout.addWidget(self.template_table)
        layout.addWidget(templates_card, 1)

        fields_card, fields_layout = self._card()
        fields_layout.addWidget(self._section_title("Campos detectados"))
        fields_layout.addWidget(self.field_table)
        layout.addWidget(fields_card, 1)

        learn_pdf_btn.clicked.connect(self.learn_pdf_template)
        load_btn.clicked.connect(self.load_template)
        list_btn.clicked.connect(self.list_pdf_fields)
        delete_btn.clicked.connect(self.delete_selected_template)
        self.template_table.cellClicked.connect(self.select_template_from_table)
        self.tabs.addTab(tab, "Plantillas")

    def _build_mapping_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Mapeo de campos",
                "Relaciona cada campo del formulario con su dato maestro correspondiente.",
            )
        )

        actions_card, actions_layout = self._card()
        actions_layout.addWidget(self._section_title("Acciones de mapeo"))
        actions = QHBoxLayout()
        create_btn = QPushButton("Crear mapeo")
        suggest_btn = QPushButton("Sugerir mapeo automatico")
        load_btn = QPushButton("Cargar mapeo")
        save_btn = QPushButton("Guardar mapeo")
        save_btn.setProperty("role", "primary")
        self.mapping_name_input = QLineEdit("mapeo_madecentro")
        actions.addWidget(create_btn)
        actions.addWidget(suggest_btn)
        actions.addWidget(load_btn)
        actions.addWidget(QLabel("Nombre"))
        actions.addWidget(self.mapping_name_input)
        actions.addWidget(save_btn)
        actions_layout.addLayout(actions)
        layout.addWidget(actions_card)

        self.mapping_table = QTableWidget(0, 2)
        self.mapping_table.setHorizontalHeaderLabels(
            ["Campo formulario", "Campo maestro"]
        )
        self._prepare_table(self.mapping_table)
        mapping_card, mapping_layout = self._card()
        mapping_layout.addWidget(self._section_title("Relaciones del formulario"))
        mapping_layout.addWidget(self.mapping_table)
        layout.addWidget(mapping_card, 1)

        create_btn.clicked.connect(self.create_empty_mapping)
        suggest_btn.clicked.connect(self.suggest_mapping)
        load_btn.clicked.connect(self.load_mapping)
        save_btn.clicked.connect(self.save_mapping)
        self.tabs.addTab(tab, "Mapeo de campos")

    def _build_fill_tab(self) -> None:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Diligenciar formulario",
                "Genera una copia utilizando la plantilla seleccionada y los datos maestros.",
            )
        )

        selection_card, selection_layout = self._card()
        selection_layout.addWidget(self._section_title("Selección actual"))
        self.selected_pdf_label = QLabel("Plantilla: sin seleccionar")
        self.selected_mapping_label = QLabel("Mapeo: sin guardar")
        self.selected_pdf_label.setObjectName("statusReady")
        self.selected_mapping_label.setObjectName("statusReady")
        selection_layout.addWidget(self.selected_pdf_label)
        selection_layout.addWidget(self.selected_mapping_label)
        layout.addWidget(selection_card)

        resources_card, resources_layout = self._card()
        resources_layout.addWidget(self._section_title("Recursos de marca"))
        resource_grid = QGridLayout()
        self.logo_status_label = QLabel()
        self.signature_status_label = QLabel()
        logo_btn = QPushButton("Abrir carpeta del logo")
        signature_btn = QPushButton("Abrir carpeta de la firma")
        resource_grid.addWidget(self.logo_status_label, 0, 0)
        resource_grid.addWidget(logo_btn, 0, 1)
        resource_grid.addWidget(self.signature_status_label, 1, 0)
        resource_grid.addWidget(signature_btn, 1, 1)
        resource_grid.setColumnStretch(0, 1)
        resources_layout.addLayout(resource_grid)
        note = QLabel(
            "Usa archivos PNG con fondo transparente. La ubicación de la firma "
            "debe definirse en cada plantilla para evitar insertarla sobre contenido."
        )
        note.setObjectName("mutedText")
        note.setWordWrap(True)
        resources_layout.addWidget(note)
        layout.addWidget(resources_card)

        action_card, action_layout = self._card()
        action_layout.addWidget(self._section_title("Generar documento"))
        actions = QHBoxLayout()
        fill_btn = QPushButton("Diligenciar formulario seleccionado")
        open_btn = QPushButton("Abrir carpeta de salida")
        fill_btn.setProperty("role", "primary")
        actions.addWidget(fill_btn)
        actions.addWidget(open_btn)
        actions.addStretch()
        action_layout.addLayout(actions)
        formats = QLabel(
            "Formatos compatibles: PDF AcroForm, Excel XLSX y Word DOCX estructurado."
        )
        formats.setObjectName("mutedText")
        action_layout.addWidget(formats)
        layout.addWidget(action_card)
        layout.addStretch()

        fill_btn.clicked.connect(self.fill_pdf)
        open_btn.clicked.connect(self.open_output_folder)
        logo_btn.clicked.connect(lambda: self._open_directory(LOGO_DIR))
        signature_btn.clicked.connect(lambda: self._open_directory(SIGNATURE_DIR))
        self._refresh_asset_status()
        self.tabs.addTab(tab, "Diligenciar formulario")

    def _build_certificate_tab(self) -> None:
        tab = QWidget()
        self.certificate_tab = tab
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(24, 22, 24, 24)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Crear certificado",
                "Busca un cliente por razón social o NIT y genera su certificado automáticamente.",
            )
        )

        search_card, search_layout = self._card()
        search_layout.addWidget(self._section_title("Buscar cliente"))
        search_actions = QHBoxLayout()
        self.certificate_search_input = QLineEdit()
        self.certificate_search_input.setPlaceholderText(
            "Ej. Maderas del Norte o 900456789"
        )
        search_btn = QPushButton("Buscar")
        new_customer_btn = QPushButton("Nuevo cliente certificado")
        search_btn.setProperty("role", "primary")
        search_actions.addWidget(self.certificate_search_input, 1)
        search_actions.addWidget(search_btn)
        search_actions.addWidget(new_customer_btn)
        search_layout.addLayout(search_actions)
        self.certificate_catalog_status = QLabel()
        self.certificate_catalog_status.setObjectName("mutedText")
        search_layout.addWidget(self.certificate_catalog_status)
        layout.addWidget(search_card)

        self.certificate_results_table = QTableWidget(0, 4)
        self.certificate_results_table.setHorizontalHeaderLabels(
            ["ID", "Razón social", "NIT", "Año de vinculación"]
        )
        self._prepare_table(self.certificate_results_table)
        results_card, results_layout = self._card()
        results_layout.addWidget(self._section_title("Clientes encontrados"))
        results_layout.addWidget(self.certificate_results_table)
        layout.addWidget(results_card, 1)

        action_card, action_layout = self._card()
        action_layout.addWidget(self._section_title("Generar certificado"))
        self.selected_certificate_label = QLabel("Cliente: sin seleccionar")
        self.selected_certificate_label.setObjectName("statusMissing")
        action_layout.addWidget(self.selected_certificate_label)
        certificate_actions = QHBoxLayout()
        create_btn = QPushButton("Crear certificado")
        open_output_btn = QPushButton("Abrir carpeta de salida")
        create_btn.setProperty("role", "primary")
        certificate_actions.addWidget(create_btn)
        certificate_actions.addWidget(open_output_btn)
        certificate_actions.addStretch()
        action_layout.addLayout(certificate_actions)
        layout.addWidget(action_card)

        search_btn.clicked.connect(self.search_certificate_customers)
        new_customer_btn.clicked.connect(self.create_certificate_customer)
        self.certificate_search_input.returnPressed.connect(
            self.search_certificate_customers
        )
        self.certificate_results_table.cellClicked.connect(
            self.select_certificate_customer
        )
        create_btn.clicked.connect(self.create_certificate)
        open_output_btn.clicked.connect(self.open_output_folder)
        self.tabs.addTab(tab, "Crear certificado")

    def _build_instructions_tab(self) -> None:
        tab = QWidget()
        outer_layout = QVBoxLayout(tab)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(24, 22, 24, 28)
        layout.setSpacing(16)
        layout.addLayout(
            self._page_intro(
                "Instructivo de uso",
                "Guía paso a paso para registrar, revisar y diligenciar formularios.",
            )
        )

        intro_card, intro_layout = self._card()
        intro_layout.addWidget(self._section_title("Antes de comenzar"))
        intro = QLabel(
            "Prepare dos versiones del mismo formulario: una completamente vacía "
            "y otra correctamente diligenciada. Ambas deben tener el mismo formato "
            "y conservar exactamente la misma estructura."
        )
        intro.setWordWrap(True)
        intro_layout.addWidget(intro)
        layout.addWidget(intro_card)

        steps = [
            (
                "1. Preparar los archivos",
                [
                    "Guarde el formulario vacío y el diligenciado en la carpeta de entrada.",
                    "Compruebe que ambos sean PDF, XLSX o DOCX del mismo tipo.",
                    "Cierre los archivos en Excel, Word o el lector de PDF antes de continuar.",
                ],
            ),
            (
                "2. Registrar la plantilla",
                [
                    "Abra la pestaña Plantillas.",
                    "Escriba un nombre fácil de identificar, por ejemplo: Vinculación de proveedores.",
                    "Pulse Registrar plantilla vacía + diligenciada.",
                    "Seleccione primero el formulario vacío y después el diligenciado.",
                    "Revise los datos detectados y confirme solamente los que sean correctos.",
                ],
            ),
            (
                "3. Revisar los datos maestros",
                [
                    "Abra la pestaña Datos maestros.",
                    "Verifique empresa, NIT, dirección, representante legal, bancos y junta directiva.",
                    "Para corregir un dato, seleccione la fila, edítela y pulse Guardar fila.",
                    "No elimine una clave si todavía se utiliza en el mapeo de una plantilla.",
                ],
            ),
            (
                "4. Revisar el mapeo",
                [
                    "Abra la pestaña Mapeo de campos.",
                    "Confirme que cada campo importante esté relacionado con la clave correcta.",
                    "Puede escribir una clave nueva cuando no exista en la lista.",
                    "Pulse Guardar mapeo antes de generar el documento.",
                ],
            ),
            (
                "5. Diligenciar el formulario",
                [
                    "Seleccione la plantilla requerida en la pestaña Plantillas.",
                    "Abra Diligenciar formulario.",
                    "Compruebe que aparecen la plantilla y el mapeo seleccionados.",
                    "Pulse Diligenciar formulario seleccionado.",
                    "Abra la carpeta de salida y revise el archivo más reciente.",
                ],
            ),
            (
                "6. Control de calidad",
                [
                    "Confirme que la fecha corresponda al día de generación.",
                    "Revise nombres, identificaciones, teléfonos, bancos y cuentas.",
                    "Compruebe todas las casillas gráficas y las selecciones marcadas con X.",
                    "Verifique especialmente miembros de junta y la sección 8.",
                    "No envíe el documento hasta completar esta revisión.",
                ],
            ),
            (
                "7. Logo y firma",
                [
                    "Coloque el logo en assets/logo, preferiblemente como logo.png.",
                    "Coloque la firma en assets/firma, preferiblemente como firma.png.",
                    "Use imágenes PNG con fondo transparente.",
                    "La firma se insertará únicamente en plantillas que tengan definida una zona de firma.",
                ],
            ),
        ]
        for step_number, (title, items) in enumerate(steps, start=1):
            card, card_layout = self._card()
            card_layout.addWidget(self._section_title(title))
            for item in items:
                label = QLabel(f"• {item}")
                label.setWordWrap(True)
                label.setObjectName("instructionStep")
                card_layout.addWidget(label)
            actions = QHBoxLayout()
            actions.addStretch()
            if step_number == 1:
                button = QPushButton("Abrir carpeta de entrada")
                button.setProperty("role", "primary")
                button.clicked.connect(
                    lambda: self._open_directory(INPUT_DIR)
                )
                actions.addWidget(button)
            elif step_number == 2:
                button = QPushButton("Ir a Plantillas")
                button.setProperty("role", "primary")
                button.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
                actions.addWidget(button)
            elif step_number == 3:
                button = QPushButton("Ir a Datos maestros")
                button.setProperty("role", "primary")
                button.clicked.connect(lambda: self.tabs.setCurrentIndex(0))
                actions.addWidget(button)
            elif step_number == 4:
                button = QPushButton("Ir a Mapeo de campos")
                button.setProperty("role", "primary")
                button.clicked.connect(lambda: self.tabs.setCurrentIndex(2))
                actions.addWidget(button)
            elif step_number == 5:
                output_button = QPushButton("Abrir carpeta de salida")
                output_button.clicked.connect(
                    lambda: self._open_directory(OUTPUT_DIR)
                )
                button = QPushButton("Ir a Diligenciar formulario")
                button.setProperty("role", "primary")
                button.clicked.connect(lambda: self.tabs.setCurrentIndex(3))
                actions.addWidget(output_button)
                actions.addWidget(button)
            elif step_number == 6:
                button = QPushButton("Abrir carpeta de salida")
                button.setProperty("role", "primary")
                button.clicked.connect(
                    lambda: self._open_directory(OUTPUT_DIR)
                )
                actions.addWidget(button)
            elif step_number == 7:
                logo_button = QPushButton("Abrir carpeta del logo")
                signature_button = QPushButton("Abrir carpeta de la firma")
                signature_button.setProperty("role", "primary")
                logo_button.clicked.connect(
                    lambda: self._open_directory(LOGO_DIR)
                )
                signature_button.clicked.connect(
                    lambda: self._open_directory(SIGNATURE_DIR)
                )
                actions.addWidget(logo_button)
                actions.addWidget(signature_button)
            card_layout.addLayout(actions)
            layout.addWidget(card)

        warning = QLabel(
            "Recomendación: conserve siempre una copia original de cada formulario "
            "vacío y no edite manualmente los archivos guardados en plantillas."
        )
        warning.setObjectName("instructionWarning")
        warning.setWordWrap(True)
        layout.addWidget(warning)
        layout.addStretch()

        scroll.setWidget(content)
        outer_layout.addWidget(scroll)
        self.tabs.addTab(tab, "Instructivo")

    def refresh_all(self) -> None:
        self.refresh_master_table()
        self.refresh_template_table()
        self.refresh_certificate_status()
        self.refresh_selected_template()
        self.refresh_mapping_combos()
        self.refresh_selected_mapping()
        self._refresh_logo()
        self._refresh_asset_status()

    def _refresh_asset_status(self) -> None:
        if not hasattr(self, "logo_status_label"):
            return
        logo = logo_path()
        signature = signature_path()
        self.logo_status_label.setText(
            f"Logo: {logo.name}" if logo else "Logo: archivo pendiente"
        )
        self.logo_status_label.setObjectName(
            "statusReady" if logo else "statusMissing"
        )
        self.signature_status_label.setText(
            f"Firma: {signature.name}"
            if signature
            else "Firma: archivo pendiente"
        )
        self.signature_status_label.setObjectName(
            "statusReady" if signature else "statusMissing"
        )
        self.logo_status_label.style().unpolish(self.logo_status_label)
        self.logo_status_label.style().polish(self.logo_status_label)
        self.signature_status_label.style().unpolish(self.signature_status_label)
        self.signature_status_label.style().polish(self.signature_status_label)

    def refresh_master_table(self) -> None:
        rows = self.db.list_master_data()
        self.master_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col, key in enumerate(["id", "clave", "valor", "categoria"]):
                item = QTableWidgetItem(str(row[key]))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.master_table.setItem(row_index, col, item)

    def refresh_template_table(self) -> None:
        certificate_ids = self._certificate_template_ids()
        rows = [
            row
            for row in self.db.list_templates()
            if int(row["id"]) not in certificate_ids
        ]
        self.template_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = [
                row["id"],
                row["nombre"],
                row["formato"].upper(),
                row["ruta_pdf"],
                row["ruta_referencia"],
            ]
            for col, value in enumerate(values):
                self.template_table.setItem(
                    row_index,
                    col,
                    QTableWidgetItem(str(value)),
                )

    def refresh_selected_template(self) -> None:
        certificate_ids = self._certificate_template_ids()
        templates = [
            item
            for item in self.db.list_templates()
            if int(item["id"]) not in certificate_ids
        ]
        if not templates:
            self._clear_template_state()
            return
        self._set_current_template(templates[0])

    def refresh_certificate_status(self) -> None:
        if not hasattr(self, "certificate_catalog_status"):
            return
        customer_count = self.db.count_certificate_customers()
        context = self._certificate_template_context()
        template_text = (
            f"Plantilla: {context[0]['nombre']}"
            if context
            else "Plantilla de certificado: pendiente"
        )
        self.certificate_catalog_status.setText(
            f"Clientes disponibles: {customer_count} | {template_text}"
        )

    def refresh_mapping_combos(self) -> None:
        master_keys = self._available_mapping_keys()
        mapped_keys = (
            list(self.form_templates.mapping(self.current_mapping_payload).values())
            if self.current_mapping_payload
            else []
        )
        available_keys = sorted(
            {key for key in master_keys + mapped_keys if key}
        )
        for row in range(self.mapping_table.rowCount()):
            combo = self.mapping_table.cellWidget(row, 1)
            if isinstance(combo, QComboBox):
                current = combo.currentText()
                combo.clear()
                combo.addItem("")
                combo.addItems(available_keys)
                combo.setCurrentText(current)

    def refresh_selected_mapping(self) -> None:
        mappings = self.db.list_mappings(self.current_template_id)
        if not mappings:
            self.current_mapping_path = None
            self.current_mapping_payload = None
            self.selected_mapping_label.setText("Mapeo: sin guardar")
            return
        latest_path = next(
            (
                Path(str(mapping["ruta_json"]))
                for mapping in mappings
                if Path(str(mapping["ruta_json"])).is_file()
            ),
            Path(str(mappings[0]["ruta_json"])),
        )
        if latest_path.exists():
            try:
                self.current_mapping_payload = self.mapping_service.load_payload(
                    latest_path
                )
            except (OSError, ValueError):
                self.current_mapping_payload = None
            self.current_mapping_path = latest_path
            self.selected_mapping_label.setText(f"Mapeo: {latest_path}")

    def import_master_data(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Importar datos maestros", "data/maestros", "Datos (*.csv *.xlsx)")
        if not file_path:
            return
        try:
            items = self._read_master_file(Path(file_path))
            if not items:
                raise ValueError("El archivo no contiene filas validas.")
            count = self.db.bulk_upsert_master_data(items)
            self._sync_master_data_file()
        except (OSError, ValueError, ValidationError) as exc:
            self._show_error("No se pudieron importar los datos", exc)
            return
        self.refresh_all()
        QMessageBox.information(self, "Importacion completa", f"Datos importados: {count}")

    def import_certificate_customers(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Importar clientes para certificados",
            str(CUSTOMER_CATALOG_PATH.parent),
            "Excel (*.xlsx)",
        )
        if not file_path:
            return
        source = Path(file_path)
        try:
            customers = self.customer_catalog.read_excel(source)
            CUSTOMER_CATALOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            if source.resolve() != CUSTOMER_CATALOG_PATH.resolve():
                shutil.copy2(source, CUSTOMER_CATALOG_PATH)
            count = self.db.replace_certificate_customers(customers)
        except (OSError, ValueError, sqlite3.DatabaseError) as exc:
            self._show_error("No se pudo importar el catálogo de clientes", exc)
            return
        self.refresh_mapping_combos()
        self.refresh_certificate_status()
        QMessageBox.information(
            self,
            "Clientes importados",
            f"Clientes disponibles para certificados: {count}",
        )

    def _load_customer_catalog_if_available(self) -> None:
        if self.db.count_certificate_customers() or not CUSTOMER_CATALOG_PATH.is_file():
            return
        try:
            customers = self.customer_catalog.read_excel(CUSTOMER_CATALOG_PATH)
            self.db.replace_certificate_customers(customers)
        except (OSError, ValueError, sqlite3.DatabaseError):
            return

    def import_completed_form(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Importar formulario diligenciado",
            "data/entrada",
            "Formularios (*.pdf *.xlsx *.docx)",
        )
        if not file_path:
            return
        source = Path(file_path)
        try:
            extracted = self.form_import.extract(source)
            resolved = self.form_import.last_resolved_payload or {}
            template_id_value = resolved.get("template_id")
            self.db.record_form_submission(
                extracted,
                source,
                int(template_id_value) if template_id_value is not None else None,
            )
            changes = self.form_import.compare(
                extracted,
                self.db.list_master_data(),
            )
        except (OSError, ValueError, RuntimeError, sqlite3.DatabaseError) as exc:
            self._show_error("No se pudo leer el formulario", exc)
            return

        if not changes:
            QMessageBox.information(
                self,
                "Sin cambios",
                "Los datos reconocidos ya coinciden con los datos maestros.",
            )
            return

        dialog = ImportPreviewDialog(changes, source.name, self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_changes()
        if not selected:
            QMessageBox.information(
                self,
                "Sin seleccion",
                "No se seleccionaron datos para actualizar.",
            )
            return
        try:
            count = self.db.apply_form_import(selected, source)
            self._sync_master_data_file()
        except (OSError, sqlite3.DatabaseError) as exc:
            self._show_error("No se pudieron actualizar los datos maestros", exc)
            return
        self.refresh_all()
        QMessageBox.information(
            self,
            "Importacion completa",
            f"Datos maestros actualizados: {count}",
        )

    def show_import_history(self) -> None:
        rows = self.db.list_import_history()
        dialog = QDialog(self)
        dialog.setWindowTitle("Historial de importaciones")
        dialog.resize(850, 420)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(rows), 5)
        table.setHorizontalHeaderLabels(
            ["ID", "Fecha", "Formato", "Cambios", "Archivo"]
        )
        table.horizontalHeader().setStretchLastSection(True)
        for row_index, row in enumerate(rows):
            values = [
                row["id"],
                row["created_at"],
                row["formato"],
                row["cantidad"],
                row["archivo"],
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(row_index, column, item)
        layout.addWidget(table)
        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()

    def _read_master_file(self, path: Path) -> list[MasterData]:
        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as file:
                reader = csv.DictReader(file)
                headers = {
                    header: header.strip().lower()
                    for header in (reader.fieldnames or [])
                    if header is not None
                }
                if {"clave", "valor"}.issubset(set(headers.values())):
                    items: list[MasterData] = []
                    for row in reader:
                        normalized = {
                            normalized_header: row.get(original_header, "")
                            for original_header, normalized_header in headers.items()
                        }
                        clave = str(normalized.get("clave") or "").strip()
                        if not clave:
                            continue
                        items.append(
                            MasterData(
                                clave=clave,
                                valor=str(normalized.get("valor") or ""),
                                categoria=str(normalized.get("categoria") or "general"),
                            )
                        )
                    return items
            with path.open("r", encoding="utf-8-sig", newline="") as file:
                rows = list(csv.reader(file))
                return [MasterData(clave=row[0], valor=row[1] if len(row) > 1 else "", categoria=row[2] if len(row) > 2 else "general") for row in rows if row]
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(cell).strip().lower() if cell is not None else "" for cell in rows[0]]
            has_headers = "clave" in headers and "valor" in headers
            data_rows = rows[1:] if has_headers else rows
            items: list[MasterData] = []
            for row in data_rows:
                if has_headers:
                    values = dict(zip(headers, row))
                    clave = values.get("clave")
                    valor = values.get("valor") or ""
                    categoria = values.get("categoria") or "general"
                else:
                    clave = row[0] if len(row) > 0 else None
                    valor = row[1] if len(row) > 1 else ""
                    categoria = row[2] if len(row) > 2 else "general"
                if clave:
                    items.append(MasterData(clave=str(clave), valor=str(valor), categoria=str(categoria)))
            return items
        finally:
            workbook.close()

    def save_master_row(self) -> None:
        try:
            item = MasterData(
                clave=self.master_key_input.text(),
                valor=self.master_value_input.text().strip(),
                categoria=self.master_category_input.text() or "general",
            )
            if self.selected_master_id is None:
                self.db.upsert_master_data(item)
            else:
                self.db.update_master_data_by_id(self.selected_master_id, item)
        except (sqlite3.IntegrityError, ValidationError) as exc:
            self._show_error("No se pudo guardar el dato maestro", exc)
            return
        self.clear_master_selection()
        self.refresh_all()
        self._sync_master_data_file()

    def load_selected_master_row(self) -> None:
        row = self._selected_master_row()
        if row is None:
            return
        self.load_master_row_from_selection(row, 0)

    def load_master_row_from_selection(self, row: int, _col: int) -> None:
        id_item = self.master_table.item(row, 0)
        key_item = self.master_table.item(row, 1)
        value_item = self.master_table.item(row, 2)
        category_item = self.master_table.item(row, 3)
        if not id_item or not key_item or not value_item or not category_item:
            return
        self.selected_master_id = int(id_item.text())
        self.master_key_input.setText(key_item.text())
        self.master_value_input.setText(value_item.text())
        self.master_category_input.setText(category_item.text())

    def edit_selected_master_row(self) -> None:
        row = self._selected_master_row()
        if row is None:
            QMessageBox.warning(self, "Sin seleccion", "Seleccione una fila de datos maestros primero.")
            return
        self.load_master_row_from_selection(row, 0)

    def delete_selected_master_row(self) -> None:
        row = self._selected_master_row()
        if row is None:
            QMessageBox.warning(self, "Sin seleccion", "Seleccione una fila de datos maestros primero.")
            return
        self.load_master_row_from_selection(row, 0)
        key = self.master_key_input.text()
        response = QMessageBox.question(
            self,
            "Confirmar eliminacion",
            f"¿Eliminar el dato maestro '{key}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if response != QMessageBox.Yes or self.selected_master_id is None:
            return
        self.db.delete_master_data_by_id(self.selected_master_id)
        self.clear_master_selection()
        self.refresh_all()
        self._sync_master_data_file()

    def clear_master_selection(self) -> None:
        self.selected_master_id = None
        self.master_key_input.clear()
        self.master_value_input.clear()
        self.master_category_input.setText("general")
        self.master_table.clearSelection()

    def _selected_master_row(self) -> int | None:
        selected_rows = self.master_table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        return selected_rows[0].row()

    def delete_selected_template(self) -> None:
        row = self._selected_template_row()
        if row is None:
            QMessageBox.warning(self, "Sin seleccion", "Seleccione una plantilla primero.")
            return
        id_item = self.template_table.item(row, 0)
        name_item = self.template_table.item(row, 1)
        if not id_item or not name_item:
            return
        template_id = int(id_item.text())
        name = name_item.text()
        response = QMessageBox.question(
            self,
            "Confirmar eliminacion",
            f"¿Eliminar la plantilla '{name}' y sus mapeos asociados?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if response != QMessageBox.Yes:
            return
        try:
            self.db.delete_template(template_id)
        except sqlite3.DatabaseError as exc:
            self._show_error("No se pudo eliminar la plantilla", exc)
            return
        if self.current_template_id == template_id:
            self._clear_template_state()
        self.refresh_all()

    def _selected_template_row(self) -> int | None:
        selected_rows = self.template_table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        return selected_rows[0].row()

    def load_template(self) -> None:
        empty_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar formulario VACIO",
            "data/entrada",
            "Formularios (*.pdf *.xlsx *.docx)",
        )
        if not empty_path:
            return
        empty_source = Path(empty_path)
        completed_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar el MISMO formulario DILIGENCIADO",
            str(empty_source.parent),
            f"Formulario (*{empty_source.suffix.lower()})",
        )
        if not completed_path:
            return
        completed_source = Path(completed_path)
        name = self.template_name_input.text().strip() or empty_source.stem
        template_id: int | None = None
        target_dir = {
            ".pdf": PDF_TEMPLATES_DIR,
            ".xlsx": EXCEL_TEMPLATES_DIR,
            ".docx": WORD_TEMPLATES_DIR,
        }.get(empty_source.suffix.lower())
        if target_dir is None:
            QMessageBox.warning(
                self,
                "Formato no soportado",
                "Use un formulario PDF, XLSX o DOCX.",
            )
            return
        version_dir = target_dir / uuid.uuid4().hex
        target = version_dir / empty_source.name
        reference = version_dir / (
            f"{completed_source.stem}_referencia{completed_source.suffix.lower()}"
        )
        try:
            payload = self.form_templates.learn(
                empty_source,
                completed_source,
                self.db.get_master_data(),
            )
            version_dir.mkdir(parents=True, exist_ok=True)
            if empty_source.resolve() != target.resolve():
                shutil.copy2(empty_source, target)
            if completed_source.resolve() != reference.resolve():
                shutil.copy2(completed_source, reference)
            template_id = self.db.add_template(
                TemplateRecord(
                    nombre=name,
                    ruta_pdf=str(target.resolve()),
                    descripcion=f"Formulario {empty_source.suffix.upper()}",
                    formato=empty_source.suffix.lower().lstrip("."),
                    ruta_referencia=str(reference.resolve()),
                )
            )
            mapping_name = f"mapeo_{name}"
            payload["template_id"] = template_id
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
            self.db.add_mapping_record(
                mapping_name,
                str(mapping_path.resolve()),
                template_id,
            )
            version = self.db.add_template_version(
                template_id,
                str(target.resolve()),
                str(reference.resolve()),
                str(mapping_path.resolve()),
                str(payload.get("template_fingerprint") or ""),
            )
            payload["template_version"] = version
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
        except (OSError, ValueError, ValidationError, sqlite3.DatabaseError) as exc:
            if template_id is not None:
                try:
                    self.db.delete_template(template_id)
                except sqlite3.DatabaseError:
                    pass
            self._show_error("No se pudo cargar la plantilla", exc)
            return
        self._set_current_template(
            {
                "id": template_id,
                "ruta_pdf": str(target.resolve()),
                "formato": empty_source.suffix.lower().lstrip("."),
                "ruta_referencia": str(reference.resolve()),
            }
        )
        self.current_mapping_path = mapping_path
        self.current_mapping_payload = payload
        self._show_current_fields()
        self._populate_mapping_table(self.form_templates.mapping(payload))
        self.refresh_template_table()
        self._offer_sample_import(payload, completed_source)
        QMessageBox.information(
            self,
            "Plantilla registrada",
            f"Formato: {empty_source.suffix.upper()}\n"
            f"Campos detectados: {len(self.form_templates.fields(payload))}",
        )

    def learn_pdf_template(self) -> None:
        pdf_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar PDF AcroForm",
            "data/entrada",
            "PDF AcroForm (*.pdf)",
        )
        if not pdf_path:
            return
        source = Path(pdf_path)
        name = self.template_name_input.text().strip() or source.stem
        try:
            payload = self.form_templates.learn_pdf_acroform(
                source,
                self.db.get_master_data(),
            )
        except (OSError, ValueError, RuntimeError) as exc:
            self._show_error("No se pudo aprender la plantilla PDF", exc)
            return

        dialog = PdfLearningDialog(
            self.form_templates.mapping(payload),
            self._available_mapping_keys(),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        mapping = dialog.mapping()
        self.form_templates.apply_field_mapping(payload, mapping)
        if set(mapping.values()) & CUSTOMER_KEYS:
            payload["purpose"] = "certificate"

        version_dir = PDF_TEMPLATES_DIR / uuid.uuid4().hex
        target = version_dir / source.name
        mapping_name = f"mapeo_{name}"
        try:
            version_dir.mkdir(parents=True, exist_ok=True)
            if source.resolve() != target.resolve():
                shutil.copy2(source, target)
            template_id = self.db.add_template(
                TemplateRecord(
                    nombre=name,
                    ruta_pdf=str(target.resolve()),
                    descripcion="Plantilla PDF AcroForm aprendida",
                    formato="pdf",
                    ruta_referencia="",
                )
            )
            payload["template_id"] = template_id
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
            self.db.add_mapping_record(
                mapping_name,
                str(mapping_path.resolve()),
                template_id,
            )
            version = self.db.add_template_version(
                template_id,
                str(target.resolve()),
                "",
                str(mapping_path.resolve()),
                str(payload.get("template_fingerprint") or ""),
            )
            payload["template_version"] = version
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
        except (OSError, ValueError, ValidationError, sqlite3.DatabaseError) as exc:
            self._show_error("No se pudo guardar la plantilla PDF", exc)
            return

        self._set_current_template(
            {
                "id": template_id,
                "ruta_pdf": str(target.resolve()),
                "formato": "pdf",
                "ruta_referencia": "",
            }
        )
        self.current_mapping_path = mapping_path
        self.current_mapping_payload = payload
        self.mapping_name_input.setText(mapping_name)
        self._show_current_fields()
        self._populate_mapping_table(mapping)
        self.refresh_all()
        if payload.get("purpose") == "certificate":
            self.tabs.setCurrentWidget(self.certificate_tab)
        QMessageBox.information(
            self,
            "Plantilla aprendida",
            f"Campos AcroForm detectados: {len(mapping)}\n"
            f"Mapeo guardado en:\n{mapping_path}",
        )

    def select_template_from_table(self, row: int, _col: int) -> None:
        path_item = self.template_table.item(row, 3)
        if not path_item:
            return
        path = Path(path_item.text())
        id_item = self.template_table.item(row, 0)
        if not id_item:
            return
        format_item = self.template_table.item(row, 2)
        reference_item = self.template_table.item(row, 4)
        self._set_current_template(
            {
                "id": int(id_item.text()),
                "ruta_pdf": str(path),
                "formato": format_item.text().lower() if format_item else path.suffix,
                "ruta_referencia": reference_item.text() if reference_item else "",
            }
        )
        try:
            self._show_current_fields()
        except (OSError, ValueError, RuntimeError) as exc:
            self._show_error("No se pudieron leer los campos", exc)

    def list_pdf_fields(self) -> None:
        if not self.current_pdf:
            QMessageBox.warning(
                self,
                "Sin plantilla",
                "Seleccione o registre una plantilla primero.",
            )
            return
        try:
            self._show_current_fields()
            fields = self.current_fields
        except (OSError, ValueError, RuntimeError) as exc:
            self._show_error("No se pudieron leer los campos", exc)
            return
        QMessageBox.information(
            self,
            "Campos detectados",
            f"Campos: {len(fields)}",
        )

    def create_empty_mapping(self) -> None:
        if not self.current_fields:
            self.list_pdf_fields()
        self._populate_mapping_table({field: "" for field in self.current_fields})

    def suggest_mapping(self) -> None:
        if not self.current_fields:
            self.list_pdf_fields()
        if self.current_mapping_payload:
            self.form_templates.suggest_mapping(
                self.current_mapping_payload,
                self.db.get_master_data(),
            )
            mapping = self.form_templates.mapping(self.current_mapping_payload)
        else:
            master_keys = list(self.db.get_master_data().keys())
            mapping = self.ai_mapping.suggest_mapping(
                self.current_fields,
                master_keys,
            )
        self._populate_mapping_table(mapping)

    def _populate_mapping_table(self, mapping: dict[str, str]) -> None:
        master_keys = sorted(
            {
                *self._available_mapping_keys(),
                *(key for key in mapping.values() if key),
            }
        )
        self.mapping_table.setRowCount(len(mapping))
        for row, (field_name, master_key) in enumerate(mapping.items()):
            field_item = QTableWidgetItem(field_name)
            field_item.setFlags(field_item.flags() & ~Qt.ItemIsEditable)
            combo = QComboBox()
            combo.setEditable(True)
            combo.addItem("")
            combo.addItems(master_keys)
            combo.setCurrentText(master_key)
            self.mapping_table.setItem(row, 0, field_item)
            self.mapping_table.setCellWidget(row, 1, combo)

    def _available_mapping_keys(self) -> list[str]:
        return sorted(
            {
                *self.db.get_master_data().keys(),
                *CUSTOMER_KEYS,
                "dia_expedicion",
                "mes_expedicion",
                "ano_expedicion",
            }
        )

    def _mapping_from_table(self) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for row in range(self.mapping_table.rowCount()):
            field_item = self.mapping_table.item(row, 0)
            combo = self.mapping_table.cellWidget(row, 1)
            if field_item and isinstance(combo, QComboBox):
                mapping[field_item.text()] = combo.currentText()
        return mapping

    def _effective_mapping(self, mapping: dict[str, str]) -> dict[str, str]:
        return {field_name: master_key for field_name, master_key in mapping.items() if master_key}

    def save_mapping(self) -> None:
        mapping = self._mapping_from_table()
        if not mapping:
            QMessageBox.warning(self, "Mapeo vacio", "Cree o sugiera un mapeo primero.")
            return
        if not self._effective_mapping(mapping):
            QMessageBox.warning(self, "Mapeo sin campos", "Seleccione al menos un campo maestro antes de guardar.")
            return
        name = self.mapping_name_input.text().strip() or "mapeo_madecentro"
        file_name = f"plantilla_{self.current_template_id}_{name}" if self.current_template_id else name
        try:
            payload = self.current_mapping_payload or {
                "format": self.current_pdf.suffix.lower().lstrip(".")
                if self.current_pdf
                else "pdf",
                "mapping": mapping,
            }
            self.form_templates.apply_field_mapping(payload, mapping)
            self.current_mapping_path = self.mapping_service.save_payload(
                payload,
                name,
                file_name,
            )
            self.current_mapping_payload = payload
            self.db.add_mapping_record(
                name,
                str(self.current_mapping_path.resolve()),
                self.current_template_id,
            )
        except (OSError, ValueError) as exc:
            self._show_error("No se pudo guardar el mapeo", exc)
            return
        self.selected_mapping_label.setText(f"Mapeo: {self.current_mapping_path}")
        if self.current_reference_path and self.current_reference_path.exists():
            self._offer_sample_import(payload, self.current_reference_path)
        QMessageBox.information(self, "Mapeo guardado", str(self.current_mapping_path))

    def fill_pdf(self) -> None:
        if not self.current_pdf:
            QMessageBox.warning(
                self,
                "Sin plantilla",
                "Seleccione una plantilla primero.",
            )
            return
        mapping = self._mapping_from_table()
        payload = self.current_mapping_payload
        if (
            (not mapping or not self._effective_mapping(mapping))
            and self.current_mapping_path
        ):
            try:
                payload = self.mapping_service.load_payload(
                    self.current_mapping_path
                )
                mapping = self.form_templates.mapping(payload)
            except (OSError, ValueError) as exc:
                self._show_error("No se pudo cargar el mapeo", exc)
                return
        mapping = self._complete_dynamic_mapping(mapping)
        effective_mapping = self._effective_mapping(mapping)
        if not effective_mapping:
            QMessageBox.warning(self, "Sin mapeo", "Cree o cargue un mapeo primero.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        if payload is None:
            payload = {
                "format": self.current_pdf.suffix.lower().lstrip("."),
                "mapping": mapping,
            }
        mapping = self._complete_dynamic_mapping(mapping)
        self.form_templates.apply_field_mapping(payload, mapping)
        output_path = OUTPUT_DIR / (
            f"{self.current_pdf.stem}_diligenciado_{timestamp}"
            f"{self.current_pdf.suffix.lower()}"
        )
        try:
            master_data = self._generation_data(mapping)
            if master_data is None:
                return
            result = self.form_templates.fill(
                self.current_pdf,
                output_path,
                payload,
                master_data,
            )
        except (OSError, ValueError, RuntimeError) as exc:
            self._show_error("No se pudo diligenciar el formulario", exc)
            return
        QMessageBox.information(self, "Formulario diligenciado", str(result))

    def _complete_dynamic_mapping(
        self,
        mapping: dict[str, str],
    ) -> dict[str, str]:
        completed = dict(mapping)
        dynamic_keys = {
            *CUSTOMER_KEYS,
            "dia_expedicion",
            "mes_expedicion",
            "ano_expedicion",
        }
        for field_name in completed:
            if not completed[field_name] and field_name in dynamic_keys:
                completed[field_name] = field_name
        return completed

    def _generation_data(
        self,
        mapping: dict[str, str],
    ) -> dict[str, str] | None:
        data = self.db.get_master_data()
        required_keys = {key for key in mapping.values() if key}
        if required_keys & CUSTOMER_KEYS:
            customer = self._select_certificate_customer()
            if customer is None:
                return None
            year = str(customer["anio_vinculacion"])
            data.update(
                {
                    "cliente_razon_social": str(customer["razon_social"]),
                    "cliente_nit": str(customer["nit"]),
                    "cliente_anio_vinculacion": year,
                    "cliente_ano_vinculacion": year,
                }
            )
        data.update(self._date_generation_data())
        return data

    def _select_certificate_customer(self) -> dict[str, object] | None:
        if self.db.count_certificate_customers() == 0:
            QMessageBox.warning(
                self,
                "Sin clientes",
                "Importe primero el Excel de clientes certificados desde "
                "Datos maestros.",
            )
            return None
        query, accepted = QInputDialog.getText(
            self,
            "Buscar cliente",
            "Razón social o NIT:",
        )
        if not accepted or not query.strip():
            return None
        matches = self.db.search_certificate_customers(query)
        if not matches:
            QMessageBox.warning(
                self,
                "Cliente no encontrado",
                "No existe un cliente que coincida con la búsqueda.",
            )
            return None
        if len(matches) == 1:
            return matches[0]
        labels = [
            f"{item['razon_social']} | {item['nit']}"
            for item in matches
        ]
        selected, accepted = QInputDialog.getItem(
            self,
            "Seleccionar cliente",
            "Coincidencias:",
            labels,
            0,
            False,
        )
        if not accepted:
            return None
        return matches[labels.index(selected)]

    def create_certificate_customer(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Nuevo cliente certificado")
        dialog.resize(520, 220)
        layout = QVBoxLayout(dialog)
        form = QGridLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        reason_input = QLineEdit()
        nit_input = QLineEdit()
        year_input = QLineEdit()
        reason_input.setPlaceholderText("Razón social")
        nit_input.setPlaceholderText("Ej. 900.456.789-1")
        year_input.setPlaceholderText("Ej. 2010")

        form.addWidget(QLabel("Razón social"), 0, 0)
        form.addWidget(reason_input, 0, 1)
        form.addWidget(QLabel("NIT"), 1, 0)
        form.addWidget(nit_input, 1, 1)
        form.addWidget(QLabel("Año de vinculación"), 2, 0)
        form.addWidget(year_input, 2, 1)
        form.setColumnStretch(1, 1)
        layout.addLayout(form)

        actions = QHBoxLayout()
        cancel_btn = QPushButton("Cancelar")
        save_btn = QPushButton("Guardar cliente")
        save_btn.setProperty("role", "primary")
        actions.addStretch()
        actions.addWidget(cancel_btn)
        actions.addWidget(save_btn)
        layout.addLayout(actions)

        saved_customer: dict[str, object] | None = None
        updated_existing_excel = False

        def save_customer() -> None:
            nonlocal saved_customer, updated_existing_excel
            reason = reason_input.text().strip()
            nit = nit_input.text().strip()
            year = year_input.text().strip()
            normalized_nit = self.customer_catalog.normalize_nit(nit)
            if not reason or not nit or not year:
                QMessageBox.warning(
                    self,
                    "Datos incompletos",
                    "Ingrese razón social, NIT y año de vinculación.",
                )
                return
            if not normalized_nit:
                QMessageBox.warning(
                    self,
                    "NIT inválido",
                    "Ingrese un NIT con al menos un número.",
                )
                return
            if not year.isdigit() or len(year) != 4:
                QMessageBox.warning(
                    self,
                    "Año inválido",
                    "Ingrese el año de vinculación con cuatro dígitos.",
                )
                return

            customer = {
                "cliente_razon_social": reason,
                "cliente_nit": nit,
                "cliente_anio_vinculacion": year,
                "cliente_ano_vinculacion": year,
                "nit_normalizado": normalized_nit,
            }
            try:
                updated_existing_excel = self.customer_catalog.upsert_excel(
                    CUSTOMER_CATALOG_PATH,
                    customer,
                )
                saved_customer = self.db.upsert_certificate_customer(customer)
            except (OSError, ValueError, sqlite3.DatabaseError) as exc:
                self._show_error("No se pudo guardar el cliente certificado", exc)
                return
            dialog.accept()

        cancel_btn.clicked.connect(dialog.reject)
        save_btn.clicked.connect(save_customer)
        if dialog.exec() != QDialog.Accepted or saved_customer is None:
            return

        self.refresh_certificate_status()
        self.certificate_search_input.setText(str(saved_customer["nit"]))
        self.search_certificate_customers()
        self._select_certificate_customer_by_id(int(saved_customer["id"]))
        action = "actualizado" if updated_existing_excel else "creado"
        QMessageBox.information(
            self,
            "Cliente guardado",
            f"Cliente certificado {action}: {saved_customer['razon_social']}",
        )

    def search_certificate_customers(self) -> None:
        query = self.certificate_search_input.text().strip()
        if not query:
            QMessageBox.warning(
                self,
                "Búsqueda vacía",
                "Escriba una razón social o un NIT.",
            )
            return
        matches = self.db.search_certificate_customers(query, limit=100)
        self.certificate_results_table.setRowCount(len(matches))
        for row_index, customer in enumerate(matches):
            values = [
                customer["id"],
                customer["razon_social"],
                customer["nit"],
                customer["anio_vinculacion"],
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.certificate_results_table.setItem(
                    row_index,
                    column,
                    item,
                )
        self.selected_certificate_customer = None
        self._set_certificate_selection_label(None)
        if not matches:
            QMessageBox.information(
                self,
                "Sin resultados",
                "No se encontraron clientes con ese nombre o NIT.",
            )

    def select_certificate_customer(self, row: int, _column: int) -> None:
        id_item = self.certificate_results_table.item(row, 0)
        name_item = self.certificate_results_table.item(row, 1)
        nit_item = self.certificate_results_table.item(row, 2)
        year_item = self.certificate_results_table.item(row, 3)
        if not all((id_item, name_item, nit_item, year_item)):
            return
        self.selected_certificate_customer = {
            "id": int(id_item.text()),
            "razon_social": name_item.text(),
            "nit": nit_item.text(),
            "anio_vinculacion": year_item.text(),
        }
        self._set_certificate_selection_label(
            self.selected_certificate_customer
        )

    def _select_certificate_customer_by_id(self, customer_id: int) -> None:
        for row in range(self.certificate_results_table.rowCount()):
            id_item = self.certificate_results_table.item(row, 0)
            if id_item and int(id_item.text()) == customer_id:
                self.certificate_results_table.selectRow(row)
                self.select_certificate_customer(row, 0)
                return

    def _set_certificate_selection_label(
        self,
        customer: dict[str, object] | None,
    ) -> None:
        if customer:
            self.selected_certificate_label.setText(
                f"Cliente: {customer['razon_social']} | "
                f"NIT {customer['nit']} | "
                f"Vinculado desde {customer['anio_vinculacion']}"
            )
            object_name = "statusReady"
        else:
            self.selected_certificate_label.setText(
                "Cliente: sin seleccionar"
            )
            object_name = "statusMissing"
        self.selected_certificate_label.setObjectName(object_name)
        self.selected_certificate_label.style().unpolish(
            self.selected_certificate_label
        )
        self.selected_certificate_label.style().polish(
            self.selected_certificate_label
        )

    def create_certificate(self) -> None:
        customer = self.selected_certificate_customer
        if customer is None:
            QMessageBox.warning(
                self,
                "Sin cliente",
                "Busque y seleccione un cliente antes de crear el certificado.",
            )
            return
        context = self._certificate_template_context()
        if context is None:
            QMessageBox.warning(
                self,
                "Sin plantilla",
                "Aprenda primero una plantilla PDF que utilice claves cliente_*.",
            )
            return
        template, payload, _mapping_path = context
        mapping = self._complete_dynamic_mapping(
            self.form_templates.mapping(payload)
        )
        self.form_templates.apply_field_mapping(payload, mapping)
        master_data = self._certificate_generation_data(customer)
        safe_name = "".join(
            char if char.isalnum() else "_"
            for char in str(customer["razon_social"])
        ).strip("_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        source = Path(str(template["ruta_pdf"]))
        output_path = OUTPUT_DIR / (
            f"Certificado_{safe_name}_{timestamp}{source.suffix.lower()}"
        )
        try:
            result = self.form_templates.fill(
                source,
                output_path,
                payload,
                master_data,
            )
        except (OSError, ValueError, RuntimeError) as exc:
            self._show_error("No se pudo crear el certificado", exc)
            return
        QMessageBox.information(
            self,
            "Certificado creado",
            str(result),
        )

    def _certificate_generation_data(
        self,
        customer: dict[str, object],
    ) -> dict[str, str]:
        year = str(customer["anio_vinculacion"])
        data = self._date_generation_data()
        data.update(
            {
                "cliente_razon_social": str(customer["razon_social"]),
                "cliente_nit": str(customer["nit"]),
                "cliente_anio_vinculacion": year,
                "cliente_ano_vinculacion": year,
            }
        )
        return data

    def _date_generation_data(self) -> dict[str, str]:
        now = datetime.now()
        months = (
            "",
            "enero",
            "febrero",
            "marzo",
            "abril",
            "mayo",
            "junio",
            "julio",
            "agosto",
            "septiembre",
            "octubre",
            "noviembre",
            "diciembre",
        )
        return {
            "dia_expedicion": str(now.day),
            "mes_expedicion": months[now.month],
            "ano_expedicion": str(now.year),
        }

    def _certificate_template_ids(self) -> set[int]:
        ids: set[int] = set()
        for template in self.db.list_templates():
            template_id = int(template["id"])
            for mapping in self.db.list_mappings(template_id):
                path = Path(str(mapping["ruta_json"]))
                if not path.is_file():
                    continue
                try:
                    payload = self.mapping_service.load_payload(path)
                except (OSError, ValueError):
                    continue
                mapped_keys = set(
                    self.form_templates.mapping(payload).values()
                )
                if (
                    payload.get("purpose") == "certificate"
                    or bool(mapped_keys & CUSTOMER_KEYS)
                ):
                    ids.add(template_id)
                    break
        return ids

    def _certificate_template_context(
        self,
    ) -> tuple[dict[str, object], dict[str, object], Path] | None:
        for template in self.db.list_templates():
            template_id = int(template["id"])
            for mapping in self.db.list_mappings(template_id):
                path = Path(str(mapping["ruta_json"]))
                if not path.is_file():
                    continue
                try:
                    payload = self.mapping_service.load_payload(path)
                except (OSError, ValueError):
                    continue
                mapped_keys = set(
                    self.form_templates.mapping(payload).values()
                )
                if (
                    payload.get("purpose") == "certificate"
                    or bool(mapped_keys & CUSTOMER_KEYS)
                ):
                    return template, payload, path
        return None

    def load_mapping(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Cargar mapeo",
            str(self.current_mapping_path.parent if self.current_mapping_path else Path("plantillas/mapeos")),
            "Mapeos JSON (*.json)",
        )
        if not file_path:
            return
        try:
            payload = self.mapping_service.load_payload(file_path)
            mapping = self.form_templates.mapping(payload)
        except (OSError, ValueError) as exc:
            self._show_error("No se pudo cargar el mapeo", exc)
            return
        self.current_mapping_path = Path(file_path)
        self.current_mapping_payload = payload
        self._populate_mapping_table(mapping)
        self.selected_mapping_label.setText(f"Mapeo: {self.current_mapping_path}")

    def _set_current_template(self, template: dict[str, object]) -> None:
        path = Path(str(template["ruta_pdf"]))
        if path.suffix.lower() not in {".pdf", ".xlsx", ".docx"} or not path.exists():
            self._clear_template_state()
            return
        self.current_template_id = int(template["id"])
        self.current_pdf = path
        reference = str(template.get("ruta_referencia") or "")
        self.current_reference_path = Path(reference) if reference else None
        self.current_fields = []
        self.current_mapping_path = None
        self.current_mapping_payload = None
        self.field_table.setRowCount(0)
        self.mapping_table.setRowCount(0)
        self.selected_pdf_label.setText(
            f"Plantilla {path.suffix.upper()}: {path}"
        )
        self.refresh_selected_mapping()
        if self.current_mapping_payload:
            mapping = self.form_templates.mapping(self.current_mapping_payload)
            self._populate_mapping_table(mapping)

    def _show_current_fields(self) -> None:
        if not self.current_pdf:
            return
        payload = self.current_mapping_payload
        if payload is None and self.current_mapping_path:
            payload = self.mapping_service.load_payload(self.current_mapping_path)
            self.current_mapping_payload = payload

        rows: list[list[str]] = []
        if payload is not None:
            fields = self.form_templates.fields(payload)
            self.current_fields = [
                str(field.get("field_id") or "")
                for field in fields
            ]
            for field in fields:
                location = str(
                    field.get("location")
                    or (
                        f"{field.get('sheet')}!{field.get('cell')}"
                        if field.get("sheet") and field.get("cell")
                        else ""
                    )
                )
                rows.append(
                    [
                        str(field.get("field_id") or ""),
                        str(
                            field.get("kind")
                            or field.get("value_type")
                            or self.current_pdf.suffix.lower().lstrip(".")
                        ),
                        location,
                        str(field.get("sample_value") or ""),
                    ]
                )
        elif self.current_pdf.suffix.lower() == ".pdf":
            pdf_fields = self.pdf_fields.list_fields(self.current_pdf)
            self.current_fields = [field.field_name for field in pdf_fields]
            rows = [
                [
                    field.field_name,
                    field.field_type,
                    f"Pagina {field.page}",
                    field.value or "",
                ]
                for field in pdf_fields
            ]
        else:
            raise ValueError("La plantilla no tiene un mapeo asociado.")

        self.field_table.setRowCount(len(rows))
        for row_index, values in enumerate(rows):
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.field_table.setItem(row_index, column, item)

    def _offer_sample_import(
        self,
        payload: dict[str, object],
        source: Path,
    ) -> None:
        extracted = self.form_templates.sample_values(payload)
        if not extracted:
            return
        changes = self.form_import.compare(
            extracted,
            self.db.list_master_data(),
        )
        if not changes:
            return
        dialog = ImportPreviewDialog(changes, source.name, self)
        dialog.setWindowTitle("Importar datos del formulario de referencia")
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_changes()
        if selected:
            self.db.apply_form_import(selected, source)
            self._sync_master_data_file()
            self.refresh_master_table()
            self.refresh_mapping_combos()

    def _sync_master_data_file(self) -> None:
        rows = self.db.list_master_data()
        MASTER_DATA_EXPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
        with MASTER_DATA_EXPORT_PATH.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["clave", "valor", "categoria"])
            writer.writeheader()
            for row in rows:
                writer.writerow(
                    {
                        "clave": row["clave"],
                        "valor": row["valor"],
                        "categoria": row["categoria"],
                    }
                )

    def _clear_template_state(self) -> None:
        self.current_template_id = None
        self.current_pdf = None
        self.current_reference_path = None
        self.current_fields = []
        self.current_mapping_path = None
        self.current_mapping_payload = None
        self.selected_pdf_label.setText("Plantilla: sin seleccionar")
        self.selected_mapping_label.setText("Mapeo: sin guardar")

    def _show_error(self, title: str, exc: Exception) -> None:
        QMessageBox.critical(self, title, str(exc))

    def _validate_mapping_for_current_pdf(self, mapping: dict[str, str]) -> None:
        if not self.current_pdf:
            raise ValueError("Seleccione una plantilla antes de cargar el mapeo.")
        valid_fields = {field.field_name for field in self.pdf_fields.list_fields(self.current_pdf)}
        unknown_fields = sorted(set(mapping) - valid_fields)
        if unknown_fields:
            preview = ", ".join(unknown_fields[:5])
            suffix = "..." if len(unknown_fields) > 5 else ""
            raise ValueError(
                f"El mapeo contiene campos que no existen en la plantilla: {preview}{suffix}"
            )

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        self._open_directory(OUTPUT_DIR)

    def _open_directory(self, directory: Path) -> None:
        directory.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(directory)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(directory)])
        else:
            subprocess.Popen(["xdg-open", str(directory)])

