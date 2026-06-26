import json
import tempfile
import unittest
from datetime import datetime
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

from app.services.ai_mapping_service import AiMappingService
from app.services.mapping_service import MappingService


class MappingServiceTests(unittest.TestCase):
    def test_load_mapping_validates_shape_and_types(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "mapping.json"
            path.write_text(json.dumps({"mapping": {"txt_nit": "nit"}}), encoding="utf-8")
            self.assertEqual(
                MappingService().load_mapping(path),
                {"txt_nit": "nit"},
            )

            path.write_text(json.dumps({"mapping": []}), encoding="utf-8")
            with self.assertRaises(ValueError):
                MappingService().load_mapping(path)

    def test_ai_mapping_uses_aliases_and_exact_names(self) -> None:
        mapping = AiMappingService(api_key="").suggest_mapping(
            ["txt_nit", "txt_email_principal", "chk_iva_si"],
            ["nit", "correo"],
        )

        self.assertEqual(mapping["txt_nit"], "nit")
        self.assertEqual(mapping["txt_email_principal"], "correo")
        self.assertEqual(mapping["chk_iva_si"], "")

    def test_ai_mapping_understands_human_excel_labels(self) -> None:
        mapping = AiMappingService(api_key="").suggest_mapping(
            [
                "NIT o No. De Identificación:",
                "Dirección Principal:",
                "Ciudad Domicilio Principal:",
                "Ventas o ingresos",
            ],
            ["nit", "direccion"],
        )

        self.assertEqual(mapping["NIT o No. De Identificación:"], "nit")
        self.assertEqual(mapping["Dirección Principal:"], "direccion")
        self.assertEqual(
            mapping["Ciudad Domicilio Principal:"],
            "ciudad",
        )
        self.assertEqual(mapping["Ventas o ingresos"], "ventas_ingresos")

    def test_ai_mapping_uses_openai_json_response(self) -> None:
        fake_response = SimpleNamespace(
            choices=[
                SimpleNamespace(
                    message=SimpleNamespace(
                        content=json.dumps(
                            {
                                "txt_nit": "nit",
                                "txt_nombre": "clave_inexistente",
                            }
                        )
                    )
                )
            ]
        )
        fake_client = SimpleNamespace(
            chat=SimpleNamespace(
                completions=SimpleNamespace(
                    create=lambda **_: fake_response,
                )
            )
        )

        with patch("openai.OpenAI", return_value=fake_client):
            mapping = AiMappingService(api_key="test-key", model="test-model").suggest_mapping(
                ["txt_nit", "txt_nombre"],
                ["nit", "razon_social"],
            )

        self.assertEqual(mapping["txt_nit"], "nit")
        self.assertEqual(mapping["txt_nombre"], "")

    def test_saves_and_reloads_pdf_learning_payload(self) -> None:
        payload = {
            "format": "pdf",
            "schema_version": 2,
            "mapping": {"txt_nit": "nit"},
            "template_fingerprint": "abc",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch(
                "app.services.mapping_service.MAPPINGS_DIR",
                Path(temp_dir),
            ):
                service = MappingService()
                path = service.save_payload(payload, "pdf aprendido")
                loaded = service.load_payload(path)

        self.assertEqual(loaded, payload)

    def test_save_payload_serializes_datetime_values(self) -> None:
        payload = {
            "format": "xlsx",
            "cells": [
                {
                    "field_id": "FORMULARIO!B1",
                    "sample_value": datetime(2026, 6, 17, 9, 30),
                }
            ],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch(
                "app.services.mapping_service.MAPPINGS_DIR",
                Path(temp_dir),
            ):
                service = MappingService()
                path = service.save_payload(payload, "excel aprendido")
                loaded = service.load_payload(path)

        self.assertEqual(
            loaded["cells"][0]["sample_value"],
            "2026-06-17T09:30:00",
        )


if __name__ == "__main__":
    unittest.main()
