from __future__ import annotations

import json
import tempfile
import unittest
from datetime import datetime
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

from app.services.ai_mapping_service import AiMappingService
from app.services.field_dictionary_service import FieldDictionaryService
from app.services.mapping_service import MappingService


class MappingServiceTests(unittest.TestCase):
    def _dictionary(
        self,
        temp_dir: str,
        json_content: object | None = None,
        csv_content: str | None = None,
    ) -> FieldDictionaryService:
        root = Path(temp_dir)
        json_path = root / "diccionario_madecentro.json"
        csv_path = root / "diccionario_madecentro.csv"
        if json_content is not None:
            json_path.write_text(
                json.dumps(json_content, ensure_ascii=False),
                encoding="utf-8",
            )
        if csv_content is not None:
            csv_path.write_text(csv_content, encoding="utf-8")
        return FieldDictionaryService(json_path=json_path, csv_path=csv_path)

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

    def test_field_dictionary_matches_exact_alias(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                [{"clave": "nit", "alias": "NIT proveedor"}],
            )

            self.assertEqual(
                dictionary.suggest_key("NIT proveedor", ["nit"]),
                "nit",
            )

    def test_field_dictionary_matches_normalized_alias_without_accents(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                [{"clave": "direccion", "etiqueta": "Dirección principal:"}],
            )

            self.assertEqual(
                dictionary.suggest_key("direccion principal", ["direccion"]),
                "direccion",
            )

    def test_field_dictionary_reads_json(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                {"campos": [{"master_key": "correo", "label": "Email principal"}]},
            )

            self.assertEqual(
                dictionary.suggest_key("email principal", ["correo"]),
                "correo",
            )

    def test_field_dictionary_reads_csv(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                csv_content="clave,alias\ntelefono,Teléfono empresa\n",
            )

            self.assertEqual(
                dictionary.suggest_key("telefono empresa", ["telefono"]),
                "telefono",
            )

    def test_field_dictionary_uses_json_as_primary_and_csv_as_complement(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                [{"clave": "nit", "alias": "Alias compartido"}],
                (
                    "clave,alias\n"
                    "correo,Alias compartido\n"
                    "correo,Correo adicional\n"
                ),
            )

            self.assertEqual(
                dictionary.suggest_key("alias compartido", ["nit", "correo"]),
                "nit",
            )
            self.assertEqual(
                dictionary.suggest_key("correo adicional", ["nit", "correo"]),
                "correo",
            )

    def test_field_dictionary_does_not_suggest_unknown_master_key(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                [{"clave": "clave_inexistente", "alias": "Campo conocido"}],
            )

            self.assertEqual(
                dictionary.suggest_key("Campo conocido", ["nit"]),
                "",
            )

    def test_ai_mapping_falls_back_when_dictionary_is_missing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = FieldDictionaryService(
                json_path=Path(temp_dir) / "no_existe.json",
                csv_path=Path(temp_dir) / "no_existe.csv",
            )
            mapping = AiMappingService(
                api_key="",
                field_dictionary=dictionary,
            ).suggest_mapping(["txt_email_principal"], ["correo"])

        self.assertEqual(mapping["txt_email_principal"], "correo")

    def test_ai_mapping_uses_field_dictionary_before_other_strategies(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            dictionary = self._dictionary(
                temp_dir,
                [{"clave": "nit", "alias": "Número tributario especial"}],
            )
            mapping = AiMappingService(
                api_key="",
                field_dictionary=dictionary,
            ).suggest_mapping(["Numero tributario especial"], ["nit"])

        self.assertEqual(mapping["Numero tributario especial"], "nit")

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

    def test_ai_mapping_can_skip_openai_even_with_api_key(self) -> None:
        with patch("openai.OpenAI") as openai_client:
            mapping = AiMappingService(
                api_key="test-key",
                model="test-model",
            ).suggest_mapping(
                ["txt_email_principal"],
                ["correo"],
                use_openai=False,
            )

        openai_client.assert_not_called()
        self.assertEqual(mapping["txt_email_principal"], "correo")

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
