from __future__ import annotations

from difflib import SequenceMatcher
import re
import unicodedata


class AiMappingService:
    """Future integration point for GPT/Gemini mapping suggestions."""

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

    def suggest_mapping(
        self,
        field_names: list[str],
        master_keys: list[str],
        min_score: float = 0.72,
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
