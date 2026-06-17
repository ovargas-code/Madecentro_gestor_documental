from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ImportedValue:
    master_key: str
    value: str
    source_field: str


@dataclass(frozen=True)
class ImportChange:
    master_key: str
    current_value: str
    new_value: str
    category: str
    source_field: str
