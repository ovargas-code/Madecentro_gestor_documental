from __future__ import annotations

import json
import os
from difflib import SequenceMatcher
import re
import unicodedata
from typing import Any

from dotenv import load_dotenv

from app.core.settings import BASE_DIR


class AiMappingService:
    """Suggest field-to-master-key mappings using OpenAI with local fallback."""

    FIELD_ALIASES = {
        "email_principal": "correo",
        "rep_legal_nombre": "representante_legal",
        "rep_legal_id": "representante_id",
        "rep_legal_cargo": "representante_cargo",
        "rep_legal_telefono": "representante_telefono",
        "rep_legal_celular": "representante_celular",
        "rep_legal_email": "representante_email",
        "cuenta_num_1": "cuenta_1",
        "cuenta_titular_1": "titular_cuenta_1",
        "cuenta_num_2": "cuenta_2",
        "cuenta_titular_2": "titular_cuenta_2",
        "firma_representante_legal": "representante_legal",
        "nombre_representante_legal": "representante_legal",
        "nombre_completo_del_representante_legal_principa_l_o_suplente": (
            "representante_legal"
        ),
        "cedula_representante_legal": "representante_id",
        "descripcion_de_actividad_principal": "descripcion_actividad_principal",
        "codigo_ciiu": "actividad_principal",
        "razon_social": "razon_social",
        "nit": "nit",
        "nit_o_no_de_identificacion": "nit",
        "dv": "nit",
        "direccion_principal": "direccion",
        "direccion_correspondencia": "direccion",
        "ciudad_domicilio_principal": "ciudad",
        "departamento_estado": "departamento",
        "telefono_1": "telefono",
        "movil_1": "celular",
        "correo_electronico": "correo",
        "pais": "pais",
        "dia_maximo_para_radicacion_de_facturas": "dia_radicacion_facturas",
        "ventas_o_ingresos": "ventas_ingresos",
        "ingresos_no_operacionales": "ingresos_no_operacionales",
        "otros_ingresos": "otros_ingresos",
        "total_ingresos": "total_ingresos",
        "activo_corriente": "activo_corriente",
        "activo_no_corriente": "activo_no_corriente",
        "pasivo_corriente": "pasivo_corriente",
        "pasivo_no_corriente": "pasivo_no_corriente",
        "patrimonio": "patrimonio",
        "total_costos": "total_costos",
        "total_gastos": "total_gastos",
        "declaracion_de_origen_de_fondos": "origen_fondos",
    }

    def __init__(self, api_key: str | None = None, model: str | None = None) -> None:
        load_dotenv(BASE_DIR / ".env")
        self.api_key = (api_key if api_key is not None else os.getenv("OPENAI_API_KEY", "")).strip()
        self.model = (model if model is not None else os.getenv("OPENAI_MODEL", "")).strip()
        if not self.model:
            self.model = "gpt-4.1-mini"

    def suggest_mapping(
        self,
        pdf_fields: list[Any],
        master_keys: list[str],
        min_score: float = 0.72,
    ) -> dict[str, str]:
        field_names = [self._field_name(field) for field in pdf_fields if self._field_name(field)]
        clean_master_keys = [str(key).strip() for key in master_keys if str(key).strip()]
        if self.api_key:
            return self._suggest_with_openai(field_names, clean_master_keys)
        return self._suggest_with_similarity(field_names, clean_master_keys, min_score)

    def _suggest_with_openai(
        self,
        field_names: list[str],
        master_keys: list[str],
    ) -> dict[str, str]:
        from openai import OpenAI

        client = OpenAI(api_key=self.api_key, timeout=30)
        response = client.chat.completions.create(
            model=self.model,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Eres un asistente que sugiere mapeos entre campos PDF "
                        "y claves maestras. Responde únicamente un JSON plano "
                        "con forma {\"campo_pdf\":\"clave_maestra\"}. Usa solo "
                        "claves maestras incluidas en la lista. Si no hay una "
                        "clave confiable, usa una cadena vacía."
                    ),
                },
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "campos_pdf": field_names,
                            "claves_maestras": master_keys,
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        )
        content = response.choices[0].message.content or "{}"
        raw_mapping = json.loads(content)
        if not isinstance(raw_mapping, dict):
            raise ValueError("OpenAI no devolvió un objeto JSON de mapeo.")

        valid_fields = set(field_names)
        valid_keys = set(master_keys)
        suggestions: dict[str, str] = {field_name: "" for field_name in field_names}
        for field_name, master_key in raw_mapping.items():
            field_text = str(field_name)
            key_text = str(master_key).strip()
            if field_text in valid_fields:
                suggestions[field_text] = key_text if key_text in valid_keys else ""
        return suggestions

    def _suggest_with_similarity(
        self,
        field_names: list[str],
        master_keys: list[str],
        min_score: float,
    ) -> dict[str, str]:
        suggestions: dict[str, str] = {}
        master_key_set = set(master_keys)
        for field_name in field_names:
            canonical_field = self._canonical_field_name(field_name)
            alias = self.FIELD_ALIASES.get(canonical_field)
            if alias:
                suggestions[field_name] = alias
                continue
            if canonical_field in master_key_set:
                suggestions[field_name] = canonical_field
                continue
            if field_name.strip().lower().startswith("chk_"):
                suggestions[field_name] = ""
                continue

            best_key = ""
            best_score = 0.0
            for master_key in master_keys:
                if not self._can_fuzzy_match(canonical_field, master_key):
                    continue
                score = self._score(canonical_field, master_key)
                if score > best_score:
                    best_key = master_key
                    best_score = score
            suggestions[field_name] = best_key if best_score >= min_score else ""
        return suggestions

    def _field_name(self, field: Any) -> str:
        if isinstance(field, str):
            return field.strip()
        if isinstance(field, dict):
            value = field.get("field_name") or field.get("field_id") or field.get("name")
            return str(value or "").strip()
        value = getattr(field, "field_name", None) or getattr(field, "name", None)
        return str(value or "").strip()

    def _score(self, left: str, right: str) -> float:
        left_norm = self._normalize(left)
        right_norm = self._normalize(right)
        if not left_norm or not right_norm:
            return 0.0
        ratio = SequenceMatcher(None, left_norm, right_norm).ratio()
        contains_bonus = 0.2 if left_norm in right_norm or right_norm in left_norm else 0.0
        return min(1.0, ratio + contains_bonus)

    def _canonical_field_name(self, value: str) -> str:
        value = self._words(value)
        for prefix in ("txt_", "chk_"):
            if value.startswith(prefix):
                value = value[len(prefix):]
                break
        return value

    def _can_fuzzy_match(self, field_name: str, master_key: str) -> bool:
        field_tokens = set(self._tokens(field_name))
        key_tokens = set(self._tokens(master_key))
        if not field_tokens or not key_tokens:
            return False
        common_tokens = field_tokens & key_tokens
        if len(common_tokens) >= 2:
            return True
        field_norm = self._normalize(field_name)
        key_norm = self._normalize(master_key)
        return field_norm == key_norm

    def _tokens(self, value: str) -> list[str]:
        tokens = self._words(value).split("_")
        ignored = {"txt", "chk", "1", "2", "3"}
        return [token for token in tokens if token and token not in ignored]

    def _normalize(self, value: str) -> str:
        return self._words(value).replace("_", "")

    def _words(self, value: str) -> str:
        normalized = unicodedata.normalize("NFKD", str(value).casefold())
        ascii_value = "".join(
            char
            for char in normalized
            if not unicodedata.combining(char)
        )
        return re.sub(r"[^a-z0-9]+", "_", ascii_value).strip("_")
