from __future__ import annotations

import csv
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from app.services.bulk_template_import_service import BulkTemplateImportService


class FakeDatabase:
    def __init__(self) -> None:
        self.templates: list[object] = []
        self.mappings: list[tuple[str, str, int]] = []
        self.versions: list[tuple[int, str, str, str, str]] = []

    def get_master_data(self) -> dict[str, str]:
        return {"nit": "900123"}

    def add_template(self, template: object) -> int:
        self.templates.append(template)
        return len(self.templates)

    def add_mapping_record(self, name: str, path: str, template_id: int) -> int:
        self.mappings.append((name, path, template_id))
        return len(self.mappings)

    def add_template_version(
        self,
        template_id: int,
        template_path: str,
        reference_path: str,
        mapping_path: str,
        fingerprint: str,
    ) -> int:
        self.versions.append(
            (template_id, template_path, reference_path, mapping_path, fingerprint)
        )
        return 1

    def delete_template(self, template_id: int) -> None:
        return None


class FakeFormTemplates:
    def __init__(self, failing_name: str = "") -> None:
        self.failing_name = failing_name
        self.learned: list[tuple[Path, Path]] = []

    def learn(
        self,
        empty: Path,
        completed: Path,
        master_data: dict[str, str],
    ) -> dict[str, object]:
        if self.failing_name and self.failing_name in empty.name:
            raise ValueError("fallo controlado")
        self.learned.append((empty, completed))
        return {
            "fields": [
                {"field_id": "txt_nit", "master_key": "nit"},
                {"field_id": "txt_nombre", "master_key": ""},
            ],
            "template_fingerprint": empty.stem,
        }

    def fields(self, payload: dict[str, object]) -> list[dict[str, object]]:
        return list(payload.get("fields", []))

    def mapping(self, payload: dict[str, object]) -> dict[str, str]:
        return {
            str(field["field_id"]): str(field.get("master_key") or "")
            for field in self.fields(payload)
        }


class FakeMappingService:
    def __init__(self, root: Path) -> None:
        self.root = root

    def save_payload(
        self,
        payload: dict[str, object],
        name: str,
        file_name: str | None = None,
    ) -> Path:
        self.root.mkdir(parents=True, exist_ok=True)
        path = self.root / f"{self._safe_file_name(file_name or name)}.json"
        path.write_text("{}", encoding="utf-8")
        return path

    def _safe_file_name(self, value: str) -> str:
        cleaned = "".join(
            char if char.isalnum() or char in ("-", "_", ".") else "_"
            for char in value.strip()
        )
        return cleaned.strip(".") or "mapeo"


class BulkTemplateImportServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def _service(self, form_templates: FakeFormTemplates | None = None) -> BulkTemplateImportService:
        return BulkTemplateImportService(
            db=FakeDatabase(),
            form_templates=form_templates or FakeFormTemplates(),
            mapping_service=FakeMappingService(self.root / "mapeos"),
            template_dirs={
                ".xlsx": self.root / "plantillas" / "excel",
                ".docx": self.root / "plantillas" / "word",
                ".pdf": self.root / "plantillas" / "pdf",
            },
        )

    def test_pairs_files_by_role_markers_format_and_name_similarity(self) -> None:
        files = [
            "Sarlaft ABC vacio.xlsx",
            "Sarlaft ABC diligenciado.xlsx",
            "Proveedor XYZ blanco.docx",
            "Proveedor XYZ lleno.docx",
            "Proveedor XYZ lleno.xlsx",
            "Sin pareja vacio.pdf",
        ]
        for name in files:
            (self.root / name).write_bytes(b"test")

        pairs, unmatched = self._service().find_pairs(self.root)

        pair_names = {(pair.empty.name, pair.completed.name) for pair in pairs}
        self.assertEqual(
            pair_names,
            {
                ("Sarlaft ABC vacio.xlsx", "Sarlaft ABC diligenciado.xlsx"),
                ("Proveedor XYZ blanco.docx", "Proveedor XYZ lleno.docx"),
            },
        )
        self.assertEqual(
            {item.path.name for item in unmatched},
            {"Proveedor XYZ lleno.xlsx", "Sin pareja vacio.pdf"},
        )

    def test_import_continues_when_one_pair_fails(self) -> None:
        for name in [
            "Correcto vacio.xlsx",
            "Correcto diligenciado.xlsx",
            "Falla vacio.xlsx",
            "Falla diligenciado.xlsx",
            "Solo vacio.docx",
        ]:
            (self.root / name).write_bytes(b"test")
        report_path = self.root / "reporte.csv"
        service = self._service(FakeFormTemplates(failing_name="Falla"))

        report = service.import_folder(self.root, report_path)

        with report.open("r", encoding="utf-8-sig", newline="") as report_file:
            rows = list(csv.DictReader(report_file))
        by_name = {row["nombre_plantilla"]: row for row in rows}
        self.assertEqual(by_name["Correcto"]["estado"], "ok")
        self.assertEqual(by_name["Correcto"]["campos_detectados"], "2")
        self.assertEqual(by_name["Correcto"]["campos_mapeados"], "1")
        self.assertEqual(by_name["Falla"]["estado"], "error")
        self.assertIn("fallo controlado", by_name["Falla"]["error"])
        self.assertEqual(by_name["Solo"]["estado"], "sin_pareja")

    def test_cli_requires_openai_key_when_flag_is_used(self) -> None:
        env = os.environ.copy()
        env["OPENAI_API_KEY"] = ""

        result = subprocess.run(
            [
                sys.executable,
                "tools/bulk_import_templates.py",
                str(self.root),
                "--use-openai",
            ],
            cwd=Path(__file__).resolve().parents[1],
            env=env,
            capture_output=True,
            text=True,
        )

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("OPENAI_API_KEY", result.stderr)


if __name__ == "__main__":
    unittest.main()
