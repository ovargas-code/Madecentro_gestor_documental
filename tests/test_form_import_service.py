import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import fitz
from openpyxl import Workbook

from app.database.database_service import DatabaseService
from app.models.import_models import ImportChange
from app.models.schemas import MasterData
from app.services.form_import_service import FormImportService


class FormImportServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.mappings = self.root / "mapeos"
        self.mappings.mkdir()
        self.service = FormImportService(self.mappings)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_extracts_pdf_using_mapping(self) -> None:
        pdf_path = self.root / "form.pdf"
        document = fitz.open()
        page = document.new_page()
        widget = fitz.Widget()
        widget.field_name = "txt_nit"
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_value = "900123"
        widget.rect = fitz.Rect(20, 20, 200, 45)
        page.add_widget(widget)
        document.save(pdf_path)
        document.close()
        (self.mappings / "pdf.json").write_text(
            json.dumps({"mapping": {"txt_nit": "nit"}}),
            encoding="utf-8",
        )

        values = self.service.extract(pdf_path)

        self.assertEqual(values[0].master_key, "nit")
        self.assertEqual(values[0].value, "900123")

    def test_extracts_excel_and_removes_fixed_label(self) -> None:
        xlsx_path = self.root / "form.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "FORMULARIO"
        sheet["B2"] = "ENTIDAD:    BANCOLOMBIA"
        workbook.save(xlsx_path)
        workbook.close()
        (self.mappings / "mapeo_formulario_excel.json").write_text(
            json.dumps(
                {
                    "cells": [
                        {
                            "sheet": "FORMULARIO",
                            "cell": "B2",
                            "master_key": "banco_1",
                            "value_format": "ENTIDAD:    {value}",
                        }
                    ]
                }
            ),
            encoding="utf-8",
        )

        values = self.service.extract(xlsx_path)

        self.assertEqual(values[0].master_key, "banco_1")
        self.assertEqual(values[0].value, "BANCOLOMBIA")

    def test_extracts_board_members_and_assigns_category(self) -> None:
        xlsx_path = self.root / "form.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "FORMULARIO"
        sheet["B59"] = "19.161.351"
        sheet["C59"] = "C.C"
        sheet["E59"] = "ANTONIO ESTEBAN GOMEZ"
        sheet["B60"] = "123"
        workbook.save(xlsx_path)
        workbook.close()
        (self.mappings / "mapeo_formulario_excel.json").write_text(
            json.dumps(
                {
                    "cells": [
                        {
                            "sheet": "FORMULARIO",
                            "cell": "B59",
                            "master_key": "junta_1_id",
                        },
                        {
                            "sheet": "FORMULARIO",
                            "cell": "C59",
                            "master_key": "junta_1_tipo_id",
                        },
                        {
                            "sheet": "FORMULARIO",
                            "cell": "E59",
                            "master_key": "junta_1_nombre",
                        },
                        {
                            "sheet": "FORMULARIO",
                            "cell": "B60",
                            "master_key": "junta_2_id",
                        },
                    ]
                }
            ),
            encoding="utf-8",
        )

        values = self.service.extract(xlsx_path)
        changes = self.service.compare(values, [])

        self.assertEqual(len(values), 3)
        self.assertEqual(values[2].value, "ANTONIO ESTEBAN GOMEZ")
        self.assertNotIn("junta_2_id", {item.master_key for item in values})
        self.assertTrue(
            all(change.category == "junta_directiva" for change in changes)
        )

    def test_extracts_word_table_labels(self) -> None:
        docx_path = self.root / "form.docx"
        document_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body><w:tbl><w:tr>
            <w:tc><w:p><w:r><w:t>NIT</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>811.028.650-1</w:t></w:r></w:p></w:tc>
          </w:tr></w:tbl></w:body>
        </w:document>"""
        with ZipFile(docx_path, "w", compression=ZIP_DEFLATED) as archive:
            archive.writestr("word/document.xml", document_xml)

        values = self.service.extract(docx_path)

        self.assertEqual(values[0].master_key, "nit")
        self.assertEqual(values[0].value, "811.028.650-1")

    def test_compare_and_database_history(self) -> None:
        db = DatabaseService(self.root / "test.db")
        db.initialize()
        db.upsert_master_data(
            MasterData(clave="nit", valor="1", categoria="empresa")
        )
        changes = [
            ImportChange(
                master_key="nit",
                current_value="1",
                new_value="2",
                category="empresa",
                source_field="Excel:FORMULARIO!B2",
            )
        ]

        count = db.apply_form_import(changes, self.root / "form.xlsx")

        self.assertEqual(count, 1)
        self.assertEqual(db.get_master_data()["nit"], "2")
        history = db.list_import_history()
        self.assertEqual(history[0]["cantidad"], 1)
        self.assertEqual(history[0]["formato"], "xlsx")


if __name__ == "__main__":
    unittest.main()
