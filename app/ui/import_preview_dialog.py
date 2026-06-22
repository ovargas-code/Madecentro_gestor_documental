from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from app.models.import_models import ImportChange


class ImportPreviewDialog(QDialog):
    def __init__(
        self,
        changes: list[ImportChange],
        source_name: str,
        parent: object | None = None,
    ) -> None:
        super().__init__(parent)
        self.changes = changes
        self.setWindowTitle("Confirmar importacion de formulario")
        self.resize(950, 520)

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                f"Archivo: {source_name}\n"
                "Desmarque cualquier dato que no deba actualizarse.\n"
                "Al aceptar, se actualiza madecentro.db y se sincroniza datos_maestros.csv."
            )
        )
        self.table = QTableWidget(len(changes), 5)
        self.table.setHorizontalHeaderLabels(
            ["Importar", "Clave", "Valor actual", "Valor nuevo", "Origen"]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(True)
        for row, change in enumerate(changes):
            selected = QTableWidgetItem()
            selected.setFlags(
                Qt.ItemIsEnabled
                | Qt.ItemIsSelectable
                | Qt.ItemIsUserCheckable
            )
            selected.setCheckState(Qt.Checked)
            self.table.setItem(row, 0, selected)
            for column, value in enumerate(
                (
                    change.master_key,
                    change.current_value,
                    change.new_value,
                    change.source_field,
                ),
                start=1,
            ):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, column, item)
        layout.addWidget(self.table)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.button(QDialogButtonBox.Ok).setText("Actualizar seleccionados")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_changes(self) -> list[ImportChange]:
        return [
            change
            for row, change in enumerate(self.changes)
            if self.table.item(row, 0).checkState() == Qt.Checked
        ]
