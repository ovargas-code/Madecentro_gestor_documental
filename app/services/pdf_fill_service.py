from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import fitz

from app.core.settings import OUTPUT_DIR


TRUE_VALUES = {"1", "true", "si", "sí", "yes", "x", "on", "checked", "marcado"}
TEXT_FONT = "Helv"
TEXT_BASE_SIZE = 9.0
TEXT_MIN_SIZE = 6.0
TEXT_COLOR = (0.0,)
TEXT_HORIZONTAL_PADDING = 3.0
TEXT_HEIGHT_RATIO = 0.9
TEXT_LINE_HEIGHT = 1.2


class PdfFillService:
    def fill_pdf(
        self,
        input_pdf: str | Path,
        output_pdf: str | Path | None,
        mapping: dict[str, str],
        master_data: dict[str, str],
    ) -> Path:
        input_pdf = Path(input_pdf).resolve()
        if not input_pdf.is_file():
            raise FileNotFoundError(f"No existe el PDF de entrada: {input_pdf}")
        if output_pdf is None:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_pdf = OUTPUT_DIR / f"{input_pdf.stem}_diligenciado.pdf"
        output_pdf = Path(output_pdf).resolve()
        if input_pdf == output_pdf:
            raise ValueError("El PDF de salida debe ser diferente al PDF de entrada.")
        output_pdf.parent.mkdir(parents=True, exist_ok=True)

        temp_output = output_pdf.with_name(
            f".{output_pdf.stem}.tmp{output_pdf.suffix or '.pdf'}"
        )
        try:
            with fitz.open(input_pdf) as doc:
                field_values = {
                    field_name: master_data.get(master_key, "")
                    for field_name, master_key in mapping.items()
                    if master_key
                }
                self.normalizar_apariencia_campos_pdf(doc, field_values)
                for page in doc:
                    for widget in page.widgets() or []:
                        field_name = widget.field_name or ""
                        if field_name not in mapping:
                            continue
                        master_key = mapping[field_name]
                        if not master_key:
                            continue
                        value = master_data.get(master_key, "")
                        self._set_widget_value(doc, widget, value)
                doc.save(temp_output, incremental=False, deflate=True, garbage=4)
            temp_output.replace(output_pdf)
        finally:
            temp_output.unlink(missing_ok=True)
        return output_pdf

    def normalizar_apariencia_campos_pdf(
        self,
        doc: fitz.Document,
        field_values: dict[str, str] | None = None,
    ) -> None:
        """Apply a consistent Helvetica appearance to every AcroForm text field."""
        field_values = field_values or {}
        self._set_acroform_default_appearance(doc)

        sizes_by_field: dict[str, float] = {}
        for page in doc:
            for widget in page.widgets() or []:
                if widget.field_type != fitz.PDF_WIDGET_TYPE_TEXT:
                    continue
                field_name = widget.field_name or ""
                value = field_values.get(
                    field_name,
                    "" if widget.field_value is None else str(widget.field_value),
                )
                calculated_size = self._calculate_text_size(widget, value)
                sizes_by_field[field_name] = min(
                    calculated_size,
                    sizes_by_field.get(field_name, TEXT_BASE_SIZE),
                )

        normalized_parents: set[int] = set()
        for page in doc:
            for widget in page.widgets() or []:
                if widget.field_type != fitz.PDF_WIDGET_TYPE_TEXT:
                    continue
                widget.text_font = TEXT_FONT
                widget.text_fontsize = sizes_by_field[widget.field_name or ""]
                widget.text_color = TEXT_COLOR
                widget.update()

                parent_type, parent_value = doc.xref_get_key(widget.xref, "Parent")
                if parent_type == "xref":
                    parent_xref = int(parent_value.split()[0])
                    if parent_xref not in normalized_parents:
                        self._set_default_appearance(doc, parent_xref, TEXT_BASE_SIZE)
                        normalized_parents.add(parent_xref)

    def _set_widget_value(self, doc: fitz.Document, widget: Any, value: str) -> None:
        field_type = widget.field_type
        if field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            self._set_checkbox_value(doc, widget, self._truthy(value))
            return
        if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
            if value:
                widget.field_value = value
                widget.update()
            return
        widget.field_value = "" if value is None else str(value)
        widget.update()

    def _truthy(self, value: str) -> bool:
        return str(value).strip().lower() in TRUE_VALUES

    def _set_checkbox_value(self, doc: fitz.Document, widget: Any, checked: bool) -> None:
        if not checked:
            widget.field_value = ""
            widget.update()
            return

        states = widget.button_states() or {}
        on_state = next(
            (state for group in states.values() for state in group if state != "Off"),
            None,
        )
        if not on_state:
            raise ValueError(f"El checkbox '{widget.field_name}' no tiene estado activo.")

        # PyMuPDF 1.26 cannot round-trip some encoded PDF names such as S#ED.
        pdf_name = f"/{on_state}"
        doc.xref_set_key(widget.xref, "AS", pdf_name)
        doc.xref_set_key(widget.xref, "V", pdf_name)
        parent_type, parent_value = doc.xref_get_key(widget.xref, "Parent")
        if parent_type == "xref":
            parent_xref = int(parent_value.split()[0])
            doc.xref_set_key(parent_xref, "V", pdf_name)

    def _calculate_text_size(self, widget: Any, value: str) -> float:
        rect = widget.rect
        available_width = max(rect.width - TEXT_HORIZONTAL_PADDING * 2, 1.0)
        line_count = max(len(value.splitlines()), 1)
        height_limit = rect.height * TEXT_HEIGHT_RATIO
        if line_count > 1:
            height_limit /= TEXT_LINE_HEIGHT * line_count

        size = min(TEXT_BASE_SIZE, height_limit)
        longest_line = max(value.splitlines() or [""], key=len)
        if longest_line:
            text_width = fitz.get_text_length(
                longest_line,
                fontname="helv",
                fontsize=TEXT_BASE_SIZE,
            )
            if text_width > available_width:
                size = min(size, TEXT_BASE_SIZE * available_width / text_width)

        clamped_size = max(TEXT_MIN_SIZE, min(TEXT_BASE_SIZE, size))
        return max(TEXT_MIN_SIZE, math.floor(clamped_size * 2) / 2)

    def _set_acroform_default_appearance(self, doc: fitz.Document) -> None:
        catalog_xref = doc.pdf_catalog()
        acroform_type, acroform_value = doc.xref_get_key(catalog_xref, "AcroForm")
        if acroform_type == "xref":
            acroform_xref = int(acroform_value.split()[0])
            self._set_default_appearance(doc, acroform_xref, TEXT_BASE_SIZE)
            doc.xref_set_key(acroform_xref, "NeedAppearances", "false")
        elif acroform_type == "dict":
            doc.xref_set_key(
                catalog_xref,
                "AcroForm/DA",
                f"(/Helv {TEXT_BASE_SIZE:g} Tf 0 g)",
            )
            doc.xref_set_key(
                catalog_xref,
                "AcroForm/NeedAppearances",
                "false",
            )

    def _set_default_appearance(
        self,
        doc: fitz.Document,
        xref: int,
        font_size: float,
    ) -> None:
        size = f"{font_size:g}"
        doc.xref_set_key(xref, "DA", f"(/Helv {size} Tf 0 g)")

