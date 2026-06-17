from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

import fitz

from app.models.schemas import PdfField


class PdfFieldService:
    def list_fields(self, pdf_path: str | Path) -> list[PdfField]:
        pdf_path = Path(pdf_path)
        if not pdf_path.is_file():
            raise FileNotFoundError(f"No existe el PDF: {pdf_path}")
        if pdf_path.suffix.lower() != ".pdf":
            raise ValueError(f"El archivo no es un PDF: {pdf_path}")
        fields: list[PdfField] = []
        with fitz.open(pdf_path) as doc:
            for page_index, page in enumerate(doc):
                for widget in page.widgets() or []:
                    field_name = widget.field_name or ""
                    if not field_name:
                        continue
                    fields.append(
                        PdfField(
                            field_name=field_name,
                            field_type=self._field_type_name(widget.field_type),
                            page=page_index + 1,
                            value=widget.field_value,
                            options=self._extract_options(widget),
                        )
                    )
        return fields

    def has_editable_fields(self, pdf_path: str | Path) -> bool:
        return bool(self.list_fields(pdf_path))

    def export_fields_to_csv(self, pdf_path: str | Path, csv_path: str | Path) -> Path:
        fields = self.list_fields(pdf_path)
        csv_path = Path(csv_path)
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=["field_name", "field_type", "page", "value", "options"])
            writer.writeheader()
            for field in fields:
                writer.writerow(
                    {
                        "field_name": field.field_name,
                        "field_type": field.field_type,
                        "page": field.page,
                        "value": field.value or "",
                        "options": ";".join(field.options),
                    }
                )
        return csv_path

    def _field_type_name(self, field_type: int) -> str:
        names: dict[int, str] = {
            fitz.PDF_WIDGET_TYPE_TEXT: "text",
            fitz.PDF_WIDGET_TYPE_CHECKBOX: "checkbox",
            fitz.PDF_WIDGET_TYPE_COMBOBOX: "combobox",
            fitz.PDF_WIDGET_TYPE_LISTBOX: "listbox",
            fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "radio",
            fitz.PDF_WIDGET_TYPE_SIGNATURE: "signature",
            fitz.PDF_WIDGET_TYPE_BUTTON: "button",
        }
        return names.get(field_type, f"unknown_{field_type}")

    def _extract_options(self, widget: Any) -> list[str]:
        choices = getattr(widget, "choice_values", None)
        if not choices:
            return []
        return [str(choice) for choice in choices]

