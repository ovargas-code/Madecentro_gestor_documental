import tempfile
import unittest
from pathlib import Path

from app.database.database_service import DatabaseService
from app.models.schemas import MasterData, TemplateRecord
from app.models.import_models import ImportedValue


class DatabaseServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db = DatabaseService(Path(self.temp_dir.name) / "test.db")
        self.db.initialize()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_bulk_upsert_updates_existing_values(self) -> None:
        self.db.bulk_upsert_master_data(
            [
                MasterData(clave="nit", valor="1", categoria="empresa"),
                MasterData(clave="ciudad", valor="Bogota"),
            ]
        )
        self.db.bulk_upsert_master_data([MasterData(clave="nit", valor="2")])

        self.assertEqual(self.db.get_master_data(), {"ciudad": "Bogota", "nit": "2"})

    def test_mapping_record_is_updated_instead_of_duplicated(self) -> None:
        template_id = self.db.add_template(
            TemplateRecord(nombre="Formulario", ruta_pdf="formulario.pdf")
        )

        first_id = self.db.add_mapping_record("principal", "uno.json", template_id)
        second_id = self.db.add_mapping_record("principal", "dos.json", template_id)

        self.assertEqual(first_id, second_id)
        mappings = self.db.list_mappings(template_id)
        self.assertEqual(len(mappings), 1)
        self.assertEqual(mappings[0]["ruta_json"], "dos.json")

    def test_template_keeps_format_and_reference(self) -> None:
        self.db.add_template(
            TemplateRecord(
                nombre="Formulario Excel",
                ruta_pdf="vacio.xlsx",
                formato="xlsx",
                ruta_referencia="diligenciado.xlsx",
            )
        )

        template = self.db.list_templates()[0]

        self.assertEqual(template["formato"], "xlsx")
        self.assertEqual(
            template["ruta_referencia"],
            "diligenciado.xlsx",
        )

    def test_template_versions_are_incremental(self) -> None:
        template_id = self.db.add_template(
            TemplateRecord(nombre="Versionado", ruta_pdf="formulario.xlsx")
        )

        first = self.db.add_template_version(
            template_id,
            "v1.xlsx",
            "r1.xlsx",
            "v1.json",
            "abc",
        )
        second = self.db.add_template_version(
            template_id,
            "v2.xlsx",
            "r2.xlsx",
            "v2.json",
            "def",
        )

        self.assertEqual((first, second), (1, 2))
        versions = self.db.list_template_versions(template_id)
        self.assertEqual([item["version"] for item in versions], [2, 1])

    def test_delete_template_removes_related_records(self) -> None:
        template_id = self.db.add_template(
            TemplateRecord(nombre="Uno", ruta_pdf="uno.pdf")
        )
        other_id = self.db.add_template(
            TemplateRecord(nombre="Dos", ruta_pdf="dos.pdf")
        )
        self.db.add_mapping_record("principal", "uno.json", template_id)
        self.db.add_mapping_record("principal", "dos.json", other_id)
        self.db.add_template_version(
            template_id,
            "uno.pdf",
            "",
            "uno.json",
            "abc",
        )

        self.db.delete_template(template_id)

        templates = self.db.list_templates()
        self.assertEqual([item["id"] for item in templates], [other_id])
        self.assertEqual(self.db.list_mappings(template_id), [])
        self.assertEqual(self.db.list_template_versions(template_id), [])
        self.assertEqual(len(self.db.list_mappings(other_id)), 1)

    def test_records_all_form_answers_without_updating_master_data(self) -> None:
        source = Path(self.temp_dir.name) / "form.xlsx"
        source.write_bytes(b"example")

        form_id = self.db.record_form_submission(
            [ImportedValue("nit", "900123", "Excel:Hoja1!B2")],
            source,
        )

        self.assertGreater(form_id, 0)
        self.assertEqual(self.db.get_master_data(), {})
        self.assertEqual(
            self.db.list_form_submissions()[0]["cantidad_respuestas"],
            1,
        )

    def test_upserts_certificate_customer_by_normalized_nit(self) -> None:
        first = self.db.upsert_certificate_customer(
            {
                "cliente_razon_social": "Cliente Uno",
                "cliente_nit": "900.123.456-1",
                "cliente_anio_vinculacion": "2010",
                "nit_normalizado": "9001234561",
            }
        )
        second = self.db.upsert_certificate_customer(
            {
                "cliente_razon_social": "Cliente Uno Actualizado",
                "cliente_nit": "9001234561",
                "cliente_anio_vinculacion": "2018",
                "nit_normalizado": "9001234561",
            }
        )

        self.assertEqual(first["id"], second["id"])
        matches = self.db.search_certificate_customers("9001234561")
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["razon_social"], "Cliente Uno Actualizado")
        self.assertEqual(matches[0]["anio_vinculacion"], "2018")


if __name__ == "__main__":
    unittest.main()
