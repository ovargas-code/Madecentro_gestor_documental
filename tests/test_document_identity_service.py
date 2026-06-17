import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from app.services.document_identity_service import DocumentIdentityService
from app.services.form_import_service import FormImportService


class DocumentIdentityServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.mappings = self.root / "mapeos"
        self.mappings.mkdir()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_excel_fingerprint_ignores_mapped_values(self) -> None:
        empty = self.root / "empty.xlsx"
        completed = self.root / "completed.xlsx"
        self._create_excel(empty, "FORM_A", "")
        self._create_excel(completed, "FORM_A", "900123")
        payload = {
            "format": "xlsx",
            "cells": [
                {
                    "field_id": "FORM_A!B1",
                    "sheet": "FORM_A",
                    "cell": "B1",
                    "master_key": "nit",
                }
            ],
        }
        service = DocumentIdentityService()

        self.assertEqual(
            service.fingerprint(empty, payload),
            service.fingerprint(completed, payload),
        )

    def test_import_selects_the_manifest_for_the_matching_sheet(self) -> None:
        source = self.root / "completed.xlsx"
        self._create_excel(source, "FORM_B", "Medellin")
        payload_a = {
            "format": "xlsx",
            "cells": [
                {
                    "sheet": "FORM_A",
                    "cell": "B1",
                    "master_key": "nit",
                }
            ],
        }
        payload_b = {
            "format": "xlsx",
            "cells": [
                {
                    "sheet": "FORM_B",
                    "cell": "B1",
                    "master_key": "ciudad",
                }
            ],
        }
        (self.mappings / "a.json").write_text(
            json.dumps(payload_a),
            encoding="utf-8",
        )
        (self.mappings / "b.json").write_text(
            json.dumps(payload_b),
            encoding="utf-8",
        )

        values = FormImportService(self.mappings).extract(source)

        self.assertEqual([(item.master_key, item.value) for item in values], [
            ("ciudad", "Medellin"),
        ])

    def _create_excel(self, path: Path, sheet_name: str, value: str) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = sheet_name
        sheet["A1"] = "Campo"
        sheet["B1"] = value
        workbook.save(path)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
