from __future__ import annotations

from pathlib import Path
from typing import Any

import fitz


class PdfSignatureService:
    def apply_signature(
        self,
        pdf_path: str | Path,
        signature_path: str | Path,
        config: dict[str, Any] | None = None,
    ) -> Path:
        pdf_path = Path(pdf_path).resolve()
        signature_path = Path(signature_path).resolve()
        config = config or {}
        temp_path = pdf_path.with_name(f".{pdf_path.stem}.signature.tmp.pdf")
        try:
            with fitz.open(pdf_path) as document:
                page_number, rect = self._location(document, config)
                document[page_number].insert_image(
                    rect,
                    filename=str(signature_path),
                    keep_proportion=True,
                    overlay=True,
                )
                self._remove_signature_placeholders(document[page_number])
                document.save(temp_path, deflate=True, garbage=4)
            temp_path.replace(pdf_path)
        finally:
            temp_path.unlink(missing_ok=True)
        return pdf_path

    def _location(
        self,
        document: fitz.Document,
        config: dict[str, Any],
    ) -> tuple[int, fitz.Rect]:
        page_number = max(
            0,
            min(int(config.get("page", document.page_count - 1)), document.page_count - 1),
        )
        raw_rect = config.get("rect")
        if isinstance(raw_rect, list) and len(raw_rect) == 4:
            return page_number, fitz.Rect(*[float(value) for value in raw_rect])
        for index, page in enumerate(document):
            for widget in page.widgets() or []:
                field_name = str(widget.field_name or "").casefold()
                if "firma" in field_name or "signature" in field_name:
                    rect = fitz.Rect(widget.rect)
                    if rect.width > 5 and rect.height > 5:
                        return index, rect
        for index, page in enumerate(document):
            matches = page.search_for("firma")
            if matches:
                label = matches[-1]
                width = float(config.get("width", 150))
                height = float(config.get("height", 55))
                x0 = min(label.x0, page.rect.width - width - 20)
                y0 = min(label.y1 + 4, page.rect.height - height - 20)
                return index, fitz.Rect(x0, y0, x0 + width, y0 + height)
        page = document[page_number]
        width = float(config.get("width", 150))
        height = float(config.get("height", 55))
        margin = 24
        return page_number, fitz.Rect(
            page.rect.width - width - margin,
            page.rect.height - height - margin,
            page.rect.width - margin,
            page.rect.height - margin,
        )

    def _remove_signature_placeholders(self, page: fitz.Page) -> None:
        for widget in list(page.widgets() or []):
            field_name = str(widget.field_name or "").casefold()
            if "firma" in field_name or "signature" in field_name:
                page.delete_widget(widget)
