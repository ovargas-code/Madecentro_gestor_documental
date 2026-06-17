from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

from app.core.settings import MAPPINGS_DIR


class MappingService:
    def create_mapping(self, field_names: list[str], master_keys: list[str]) -> dict[str, str]:
        mapping: dict[str, str] = {}
        normalized_master = {self._normalize(key): key for key in master_keys}
        for field_name in field_names:
            normalized_field = self._normalize(field_name)
            mapping[field_name] = normalized_master.get(normalized_field, "")
        return mapping

    def save_mapping(
        self,
        mapping: dict[str, str],
        name: str,
        file_name: str | None = None,
    ) -> Path:
        MAPPINGS_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = self._safe_file_name(file_name or name).removesuffix(".json") + ".json"
        path = MAPPINGS_DIR / safe_name
        payload = {"name": name, "mapping": mapping}
        temp_path = path.with_suffix(f"{path.suffix}.tmp")
        try:
            temp_path.write_text(
                json.dumps(
                    payload,
                    ensure_ascii=False,
                    indent=2,
                    default=self._json_default,
                ),
                encoding="utf-8",
            )
            temp_path.replace(path)
        finally:
            temp_path.unlink(missing_ok=True)
        return path

    def save_payload(
        self,
        payload: dict[str, Any],
        name: str,
        file_name: str | None = None,
    ) -> Path:
        MAPPINGS_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = self._safe_file_name(file_name or name).removesuffix(".json") + ".json"
        path = MAPPINGS_DIR / safe_name
        temp_path = path.with_suffix(f"{path.suffix}.tmp")
        try:
            temp_path.write_text(
                json.dumps(
                    payload,
                    ensure_ascii=False,
                    indent=2,
                    default=self._json_default,
                ),
                encoding="utf-8",
            )
            temp_path.replace(path)
        finally:
            temp_path.unlink(missing_ok=True)
        return path

    def load_payload(self, path: str | Path) -> dict[str, Any]:
        payload: Any = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("El archivo de mapeo debe contener un objeto JSON.")
        return payload

    def load_mapping(self, path: str | Path) -> dict[str, str]:
        mapping_path = Path(path)
        payload: Any = json.loads(mapping_path.read_text(encoding="utf-8"))
        if isinstance(payload, dict) and "mapping" in payload:
            payload = payload["mapping"]
        if not isinstance(payload, dict):
            raise ValueError(f"El archivo no contiene un mapeo valido: {mapping_path}")
        mapping: dict[str, str] = {}
        for field_name, master_key in payload.items():
            if not isinstance(field_name, str) or not isinstance(master_key, str):
                raise ValueError(f"El mapeo contiene claves o valores invalidos: {mapping_path}")
            mapping[field_name] = master_key
        return mapping

    def _normalize(self, value: str) -> str:
        return "".join(char.lower() for char in value if char.isalnum())

    def _safe_file_name(self, value: str) -> str:
        cleaned = "".join(char if char.isalnum() or char in ("-", "_", ".") else "_" for char in value.strip())
        cleaned = cleaned.strip(".")
        return cleaned or "mapeo"

    def _json_default(self, value: object) -> str:
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        raise TypeError(
            f"Object of type {value.__class__.__name__} is not JSON serializable"
        )

