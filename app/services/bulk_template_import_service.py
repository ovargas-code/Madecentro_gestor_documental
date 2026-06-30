from __future__ import annotations

import csv
import shutil
import sqlite3
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from app.core.settings import EXCEL_TEMPLATES_DIR, PDF_TEMPLATES_DIR, WORD_TEMPLATES_DIR
from app.database.database_service import DatabaseService
from app.models.schemas import TemplateRecord
from app.services.form_template_service import FormTemplateService
from app.services.mapping_service import MappingService


EMPTY_MARKERS = (
    "sin diligenciar",
    "vacio",
    "blanco",
    "plantilla",
)
COMPLETED_MARKERS = (
    "diligenciado",
    "lleno",
    "completado",
    "referencia",
)
SUPPORTED_SUFFIXES = {".xlsx", ".docx", ".pdf"}


@dataclass(frozen=True)
class BulkTemplateCandidate:
    path: Path
    role: str
    suffix: str
    template_key: str
    template_name: str


@dataclass(frozen=True)
class BulkTemplatePair:
    empty: Path
    completed: Path
    template_name: str
    suffix: str
    score: float


class BulkTemplateImportService:
    def __init__(
        self,
        db: DatabaseService | None = None,
        form_templates: FormTemplateService | None = None,
        mapping_service: MappingService | None = None,
        template_dirs: dict[str, Path] | None = None,
    ) -> None:
        self.db = db or DatabaseService()
        self.form_templates = form_templates or FormTemplateService()
        self.mapping_service = mapping_service or MappingService()
        self.template_dirs = template_dirs or {
            ".xlsx": EXCEL_TEMPLATES_DIR,
            ".docx": WORD_TEMPLATES_DIR,
            ".pdf": PDF_TEMPLATES_DIR,
        }

    def import_folder(self, folder: str | Path, report_path: str | Path | None = None) -> Path:
        root = Path(folder)
        if not root.is_dir():
            raise ValueError(f"La carpeta no existe: {root}")
        pairs, unmatched = self.find_pairs(root)
        report = Path(report_path) if report_path else self._default_report_path(root)
        rows: list[dict[str, str]] = []
        for pair in pairs:
            rows.append(self._import_pair(pair))
        for candidate in unmatched:
            rows.append(
                self._report_row(
                    empty=candidate.path if candidate.role == "empty" else None,
                    completed=candidate.path if candidate.role == "completed" else None,
                    template_name=candidate.template_name,
                    suffix=candidate.suffix,
                    status="sin_pareja",
                )
            )
        self._write_report(report, rows)
        return report

    def find_pairs(self, folder: str | Path) -> tuple[list[BulkTemplatePair], list[BulkTemplateCandidate]]:
        candidates = self._scan_candidates(Path(folder))
        empty = [item for item in candidates if item.role == "empty"]
        completed = [item for item in candidates if item.role == "completed"]
        pair_options: list[tuple[float, BulkTemplateCandidate, BulkTemplateCandidate]] = []
        for empty_item in empty:
            for completed_item in completed:
                if empty_item.suffix != completed_item.suffix:
                    continue
                score = self._similarity(empty_item.template_key, completed_item.template_key)
                if score >= 0.55:
                    pair_options.append((score, empty_item, completed_item))
        pair_options.sort(key=lambda item: item[0], reverse=True)

        used_empty: set[Path] = set()
        used_completed: set[Path] = set()
        pairs: list[BulkTemplatePair] = []
        for score, empty_item, completed_item in pair_options:
            if empty_item.path in used_empty or completed_item.path in used_completed:
                continue
            used_empty.add(empty_item.path)
            used_completed.add(completed_item.path)
            pairs.append(
                BulkTemplatePair(
                    empty=empty_item.path,
                    completed=completed_item.path,
                    template_name=empty_item.template_name,
                    suffix=empty_item.suffix,
                    score=score,
                )
            )

        unmatched = [
            item
            for item in candidates
            if item.path not in used_empty and item.path not in used_completed
        ]
        pairs.sort(key=lambda item: (item.suffix, item.template_name.casefold()))
        unmatched.sort(key=lambda item: str(item.path).casefold())
        return pairs, unmatched

    def _scan_candidates(self, folder: Path) -> list[BulkTemplateCandidate]:
        candidates: list[BulkTemplateCandidate] = []
        for path in sorted(folder.rglob("*")):
            if not path.is_file() or path.suffix.lower() not in SUPPORTED_SUFFIXES:
                continue
            role = self._classify_role(path.stem)
            if not role:
                continue
            template_name = self._template_name(path.stem)
            candidates.append(
                BulkTemplateCandidate(
                    path=path,
                    role=role,
                    suffix=path.suffix.lower(),
                    template_key=self._normalize_name(template_name),
                    template_name=template_name,
                )
            )
        return candidates

    def _classify_role(self, stem: str) -> str | None:
        normalized = self._normalize_name(stem)
        has_empty = any(marker in normalized for marker in EMPTY_MARKERS)
        has_completed = any(marker in normalized for marker in COMPLETED_MARKERS)
        if has_empty and not has_completed:
            return "empty"
        if has_completed and not has_empty:
            return "completed"
        return None

    def _template_name(self, stem: str) -> str:
        normalized_stem = self._strip_role_markers(
            stem.replace("_", " ").replace("-", " ")
        )
        cleaned = " ".join(normalized_stem.replace("_", " ").replace("-", " ").split())
        return cleaned or stem.strip()

    def _strip_role_markers(self, value: str) -> str:
        result = value
        for marker in (*EMPTY_MARKERS, *COMPLETED_MARKERS):
            result = self._remove_marker(result, marker)
        return result

    def _remove_marker(self, value: str, marker: str) -> str:
        words = marker.split()
        parts = value.split()
        if len(words) == 1:
            return " ".join(part for part in parts if self._normalize_name(part) != marker)
        normalized_parts = [self._normalize_name(part) for part in parts]
        output: list[str] = []
        index = 0
        while index < len(parts):
            if normalized_parts[index : index + len(words)] == list(words):
                index += len(words)
                continue
            output.append(parts[index])
            index += 1
        return " ".join(output)

    def _similarity(self, left: str, right: str) -> float:
        if not left or not right:
            return 0.0
        if left == right:
            return 1.0
        return SequenceMatcher(None, left, right).ratio()

    def _normalize_name(self, value: str) -> str:
        decomposed = unicodedata.normalize("NFKD", value)
        ascii_text = "".join(char for char in decomposed if not unicodedata.combining(char))
        lowered = ascii_text.casefold()
        return " ".join(
            "".join(char if char.isalnum() else " " for char in lowered).split()
        )

    def _import_pair(self, pair: BulkTemplatePair) -> dict[str, str]:
        template_id: int | None = None
        try:
            payload = self.form_templates.learn(
                pair.empty,
                pair.completed,
                self.db.get_master_data(),
            )
            target_dir = self._template_storage_dir(self.template_dirs[pair.suffix], pair.template_name)
            target = target_dir / pair.empty.name
            reference = target_dir / f"{pair.completed.stem}_referencia{pair.completed.suffix.lower()}"
            target_dir.mkdir(parents=True, exist_ok=True)
            if pair.empty.resolve() != target.resolve():
                shutil.copy2(pair.empty, target)
            if pair.completed.resolve() != reference.resolve():
                shutil.copy2(pair.completed, reference)

            template_id = self.db.add_template(
                TemplateRecord(
                    nombre=pair.template_name,
                    ruta_pdf=str(target.resolve()),
                    descripcion=f"Formulario {pair.suffix.upper()}",
                    formato=pair.suffix.lstrip("."),
                    ruta_referencia=str(reference.resolve()),
                )
            )
            mapping_name = f"mapeo_{pair.template_name}"
            payload["template_id"] = template_id
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
            self.db.add_mapping_record(
                mapping_name,
                str(mapping_path.resolve()),
                template_id,
            )
            version = self.db.add_template_version(
                template_id,
                str(target.resolve()),
                str(reference.resolve()),
                str(mapping_path.resolve()),
                str(payload.get("template_fingerprint") or ""),
            )
            payload["template_version"] = version
            mapping_path = self.mapping_service.save_payload(
                payload,
                mapping_name,
                f"plantilla_{template_id}_{mapping_name}",
            )
            mapped, unmapped = self._mapping_counts(payload)
            return self._report_row(
                empty=pair.empty,
                completed=pair.completed,
                template_name=pair.template_name,
                suffix=pair.suffix,
                status="ok",
                detected=len(self.form_templates.fields(payload)),
                mapped=mapped,
                unmapped=unmapped,
                mapping_path=mapping_path,
            )
        except Exception as exc:
            if template_id is not None:
                try:
                    self.db.delete_template(template_id)
                except sqlite3.DatabaseError:
                    pass
            return self._report_row(
                empty=pair.empty,
                completed=pair.completed,
                template_name=pair.template_name,
                suffix=pair.suffix,
                status="error",
                error=str(exc),
            )

    def _mapping_counts(self, payload: dict[str, Any]) -> tuple[int, int]:
        mapping = self.form_templates.mapping(payload)
        mapped = sum(1 for value in mapping.values() if value)
        return mapped, len(mapping) - mapped

    def _template_storage_dir(self, root: Path, name: str) -> Path:
        safe_name = self.mapping_service._safe_file_name(name)
        return root / safe_name

    def _default_report_path(self, folder: Path) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return folder / f"reporte_importacion_plantillas_{timestamp}.csv"

    def _report_row(
        self,
        *,
        empty: Path | None,
        completed: Path | None,
        template_name: str,
        suffix: str,
        status: str,
        detected: int = 0,
        mapped: int = 0,
        unmapped: int = 0,
        error: str = "",
        mapping_path: Path | None = None,
    ) -> dict[str, str]:
        return {
            "archivo_vacio": str(empty or ""),
            "archivo_diligenciado": str(completed or ""),
            "nombre_plantilla": template_name,
            "formato": suffix.lstrip("."),
            "estado": status,
            "campos_detectados": str(detected),
            "campos_mapeados": str(mapped),
            "campos_sin_mapeo": str(unmapped),
            "error": error,
            "ruta_mapeo": str(mapping_path.resolve()) if mapping_path else "",
        }

    def _write_report(self, path: Path, rows: list[dict[str, str]]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        fieldnames = [
            "archivo_vacio",
            "archivo_diligenciado",
            "nombre_plantilla",
            "formato",
            "estado",
            "campos_detectados",
            "campos_mapeados",
            "campos_sin_mapeo",
            "error",
            "ruta_mapeo",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as report_file:
            writer = csv.DictWriter(report_file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
