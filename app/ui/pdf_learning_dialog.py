from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class PdfLearningDialog(QDialog):
    def __init__(
        self,
        mapping: dict[str, str],
        master_keys: list[str],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Aprender plantilla PDF")
        self.resize(760, 560)

        layout = QVBoxLayout(self)
        description = QLabel(
            "Relaciona cada campo AcroForm con la clave maestra que debe "
            "utilizarse al diligenciar el PDF."
        )
        description.setWordWrap(True)
        layout.addWidget(description)

        self.table = QTableWidget(len(mapping), 2)
        self.table.setHorizontalHeaderLabels(["Campo PDF", "Clave maestra"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        available_keys = sorted({*master_keys, *(key for key in mapping.values() if key)})
        for row, (field_name, master_key) in enumerate(mapping.items()):
            field_item = QTableWidgetItem(field_name)
            field_item.setFlags(field_item.flags() & ~Qt.ItemIsEditable)
            combo = QComboBox()
            combo.setEditable(True)
            combo.addItem("")
            combo.addItems(available_keys)
            combo.setCurrentText(master_key)
            self.table.setItem(row, 0, field_item)
            self.table.setCellWidget(row, 1, combo)
        self.table.resizeColumnsToContents()
        layout.addWidget(self.table, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel
        )
        buttons.button(QDialogButtonBox.Save).setText("Guardar plantilla")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def mapping(self) -> dict[str, str]:
        result: dict[str, str] = {}
        for row in range(self.table.rowCount()):
            field_item = self.table.item(row, 0)
            combo = self.table.cellWidget(row, 1)
            if field_item and isinstance(combo, QComboBox):
                result[field_item.text()] = combo.currentText().strip()
        return result
