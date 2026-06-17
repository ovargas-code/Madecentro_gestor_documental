import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from app.database.database_service import DatabaseService
from app.services.customer_catalog_service import CustomerCatalogService


class CustomerCatalogServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_reads_and_searches_customers_by_name_or_nit(self) -> None:
        source = self.root / "clientes.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "cliente_razon_social",
                "cliente_nit",
                "cliente_anio_vinculacion",
            ]
        )
        sheet.append(["MADERAS DEL NORTE S.A.S.", "900.456.789-1", 2010])
        workbook.save(source)
        workbook.close()

        customers = CustomerCatalogService().read_excel(source)
        db = DatabaseService(self.root / "test.db")
        db.initialize()
        db.replace_certificate_customers(customers)

        by_name = db.search_certificate_customers("maderas del norte")
        by_nit = db.search_certificate_customers("9004567891")
        self.assertEqual(by_name[0]["nit"], "900.456.789-1")
        self.assertEqual(by_nit[0]["anio_vinculacion"], "2010")
        self.assertEqual(
            customers[0]["cliente_ano_vinculacion"],
            "2010",
        )

    def test_rejects_catalog_without_required_columns(self) -> None:
        source = self.root / "invalid.xlsx"
        workbook = Workbook()
        workbook.active.append(["nombre", "nit"])
        workbook.save(source)
        workbook.close()

        with self.assertRaises(ValueError):
            CustomerCatalogService().read_excel(source)

    def test_upserts_customer_catalog_excel(self) -> None:
        source = self.root / "clientes.xlsx"
        service = CustomerCatalogService()
        created = service.upsert_excel(
            source,
            {
                "cliente_razon_social": "Cliente Nuevo",
                "cliente_nit": "900.111.222-3",
                "cliente_anio_vinculacion": "2020",
                "nit_normalizado": "9001112223",
            },
        )
        updated = service.upsert_excel(
            source,
            {
                "cliente_razon_social": "Cliente Nuevo S.A.S.",
                "cliente_nit": "9001112223",
                "cliente_anio_vinculacion": "2021",
                "nit_normalizado": "9001112223",
            },
        )

        customers = service.read_excel(source)

        self.assertFalse(created)
        self.assertTrue(updated)
        self.assertEqual(len(customers), 1)
        self.assertEqual(customers[0]["cliente_razon_social"], "Cliente Nuevo S.A.S.")
        self.assertEqual(customers[0]["cliente_anio_vinculacion"], "2021")


if __name__ == "__main__":
    unittest.main()
