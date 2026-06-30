from __future__ import annotations

from pathlib import Path

import fitz

from app.plantilla_generator.replacements import ReplacementRule


class PdfCleaner:
    def clean(
        self,
        source: str | Path,
        destination: str | Path,
        rules: list[ReplacementRule],
    ) -> tuple[int, list[str]]:
        document = fitz.open(source)
        total = 0
        values: list[str] = []
        try:
            insertions: list[tuple[int, fitz.Rect, str]] = []
            for page_number, page in enumerate(document):
                for rule in rules:
                    areas = page.search_for(rule.value)
                    for area in areas:
                        page.add_redact_annot(area, fill=(1, 1, 1))
                        if rule.replacement:
                            insertions.append((page_number, area, rule.replacement))
                        total += 1
                        values.append(rule.value)
                page.apply_redactions()

            for page_number, area, replacement in insertions:
                page = document[page_number]
                try:
                    page.insert_textbox(
                        area,
                        replacement,
                        fontsize=max(6, min(9, area.height * 0.8)),
                        color=(0, 0, 0),
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                except Exception:
                    pass
            document.save(destination, garbage=4, deflate=True)
        finally:
            document.close()
        return total, sorted(set(values))
