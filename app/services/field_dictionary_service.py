from __future__ import annotations

import csv
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable

from app.core.settings import MASTER_DIR


class FieldDictionaryService:
    """Local alias dictionary for recognizing form fields."""

    KEY_FIELDS = ("clave", "master_key", "campo_maestro", "key")
    ALIAS_FIELDS = (
        "alias",
        "aliases",
        "etiqueta",
        "label",
        "texto",
        "campo",
        "field",
        "nombre",
        "categoria",
    )
    DEFAULT_KEY_ALIASES = {
        "correo_electronico": "correo",
        "email": "correo",
        "e_mail": "correo",
        "identificacion_representante": "representante_id",
        "cedula_representante": "representante_id",
        "representante_legal": "representante_legal",
        "razon_social": "razon_social",
        "nit": "nit",
        "direccion": "direccion",
        "telefono": "telefono",
        "ciudad": "ciudad",
        "departamento": "departamento",
        "pais": "pais",
    }

    def __init__(
        self,
        json_path: str | Path | None = None,
        csv_path: str | Path | None = None,
    ) -> None:
        self.json_path = Path(json_path) if json_path else MASTER_DIR / "diccionario_madecentro.json"
        self.csv_path = Path(csv_path) if csv_path else MASTER_DIR / "diccionario_madecentro.csv"
        self._aliases: dict[str, str] | None = None

    def suggest_key(self, label: str, available_master_keys: list[str]) -> str:
        available = [str(key).strip() for key in available_master_keys if str(key).strip()]
        available_by_normalized = {self.normalize(key): key for key in available}
        normalized_label = self.normalize(label)
        if not normalized_label:
            return ""

        suggested = self.aliases.get(normalized_label, "")
        if not suggested:
            return ""
        if suggested in available:
            return suggested
        return available_by_normalized.get(self.normalize(suggested), "")

    @property
    def aliases(self) -> dict[str, str]:
        if self._aliases is None:
            self._aliases = self._load_aliases()
        return self._aliases

    def _load_aliases(self) -> dict[str, str]:
        aliases: dict[str, str] = {}
        if self.json_path.is_file():
            self._merge_rows(aliases, self._read_json_rows(self.json_path))
        if self.csv_path.is_file():
            self._merge_rows(aliases, self._read_csv_rows(self.csv_path))
        return aliases

    def _merge_rows(
        self,
        aliases: dict[str, str],
        rows: Iterable[dict[str, Any]],
    ) -> None:
        for row in rows:
            key = self._row_key(row)
            row_aliases = list(self._row_aliases(row))
            if not key and row_aliases:
                key = self._infer_key(row_aliases[0])
            if not key:
                continue
            for alias in row_aliases:
                normalized_alias = self.normalize(alias)
                if normalized_alias:
                    aliases.setdefault(normalized_alias, key)

    def _read_json_rows(self, path: Path) -> list[dict[str, Any]]:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
        return list(self._extract_rows(payload))

    def _read_csv_rows(self, path: Path) -> list[dict[str, Any]]:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return [dict(row) for row in csv.DictReader(handle)]

    def _extract_rows(self, payload: Any) -> Iterable[dict[str, Any]]:
        if isinstance(payload, list):
            for item in payload:
                yield from self._extract_rows(item)
        elif isinstance(payload, dict):
            if any(field in payload for field in (*self.KEY_FIELDS, *self.ALIAS_FIELDS)):
                yield payload
            else:
                for value in payload.values():
                    yield from self._extract_rows(value)

    def _row_key(self, row: dict[str, Any]) -> str:
        normalized_row = {self.normalize(key): value for key, value in row.items()}
        for field in self.KEY_FIELDS:
            value = normalized_row.get(self.normalize(field))
            if value not in (None, ""):
                return str(value).strip()
        return ""

    def _row_aliases(self, row: dict[str, Any]) -> Iterable[str]:
        normalized_row = {self.normalize(key): value for key, value in row.items()}
        for field in self.ALIAS_FIELDS:
            value = normalized_row.get(self.normalize(field))
            yield from self._values(value)

    def _values(self, value: Any) -> Iterable[str]:
        if value in (None, ""):
            return
        if isinstance(value, (list, tuple, set)):
            for item in value:
                yield from self._values(item)
            return
        text = str(value).strip()
        if text:
            yield text

    def _infer_key(self, alias: str) -> str:
        normalized = self.normalize(alias)
        return self.DEFAULT_KEY_ALIASES.get(normalized, normalized.replace(" ", "_"))

    def normalize(self, value: Any) -> str:
        normalized = unicodedata.normalize("NFKD", str(value or "").casefold())
        ascii_value = "".join(
            char
            for char in normalized
            if not unicodedata.combining(char)
        )
        text = re.sub(r"[^a-z0-9]+", " ", ascii_value)
        return " ".join(text.split())
