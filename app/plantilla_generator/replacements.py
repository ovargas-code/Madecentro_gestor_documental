from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any


DEFAULT_MARKERS = {
    "razon_social": "<<RAZON_SOCIAL>>",
    "nit": "<<NIT>>",
    "representante_legal": "<<REPRESENTANTE_LEGAL>>",
    "cedula_representante": "<<CEDULA_REPRESENTANTE>>",
    "direccion": "<<DIRECCION>>",
    "telefono": "<<TELEFONO>>",
    "correo": "<<CORREO>>",
    "banco": "<<BANCO>>",
    "cuenta_bancaria": "<<CUENTA_BANCARIA>>",
}

CATEGORY_ALIASES = {
    "razon social": "razon_social",
    "nit": "nit",
    "representante legal": "representante_legal",
    "identificacion representante": "cedula_representante",
    "cedula representante": "cedula_representante",
    "cedula representante legal": "cedula_representante",
    "direccion": "direccion",
    "telefono": "telefono",
    "correo": "correo",
    "correo electronico": "correo",
    "email": "correo",
    "banco": "banco",
    "cuenta bancaria": "cuenta_bancaria",
    "cuenta": "cuenta_bancaria",
}


@dataclass(frozen=True)
class ReplacementRule:
    category: str
    value: str
    replacement: str
    confidence: str = ""


def normalize_text(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or "").casefold())
    ascii_value = "".join(
        char
        for char in normalized
        if not unicodedata.combining(char)
    )
    text = re.sub(r"[^a-z0-9]+", " ", ascii_value)
    return " ".join(text.split())


def canonical_category(value: Any) -> str:
    normalized = normalize_text(value)
    if normalized in CATEGORY_ALIASES:
        return CATEGORY_ALIASES[normalized]
    return normalized.replace(" ", "_")


def marker_for_category(category: str) -> str:
    canonical = canonical_category(category)
    return DEFAULT_MARKERS.get(canonical, f"<<{canonical.upper()}>>")


def build_rule(
    category: Any,
    value: Any,
    replacement: Any = "",
    confidence: Any = "",
    mode: str = "markers",
) -> ReplacementRule | None:
    text_value = str(value or "").strip()
    if not text_value:
        return None
    canonical = canonical_category(category)
    if mode == "blank":
        output = ""
    else:
        output = str(replacement or "").strip() or marker_for_category(canonical)
    return ReplacementRule(
        category=canonical,
        value=text_value,
        replacement=output,
        confidence=str(confidence or "").strip(),
    )


def apply_text_replacements(
    text: str,
    rules: list[ReplacementRule],
) -> tuple[str, int, list[str]]:
    result = text
    count = 0
    values: list[str] = []
    for rule in rules:
        occurrences = result.count(rule.value)
        if not occurrences:
            continue
        result = result.replace(rule.value, rule.replacement)
        count += occurrences
        values.append(rule.value)
    return result, count, values
