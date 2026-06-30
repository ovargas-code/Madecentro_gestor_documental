from __future__ import annotations

from pathlib import Path

from app.plantilla_generator.dictionary_loader import DictionaryLoader
from app.plantilla_generator.excel_cleaner import ExcelCleaner
from app.plantilla_generator.pdf_cleaner import PdfCleaner
from app.plantilla_generator.replacements import ReplacementRule
from app.plantilla_generator.report_writer import GenerationReportRow, ReportWriter
from app.plantilla_generator.word_cleaner import WordCleaner


class TemplateGenerator:
    SUPPORTED_SUFFIXES = {".docx", ".xlsx", ".xlsm", ".pdf"}

    def __init__(self) -> None:
        self.dictionary_loader = DictionaryLoader()
        self.word_cleaner = WordCleaner()
        self.excel_cleaner = ExcelCleaner()
        self.pdf_cleaner = PdfCleaner()
        self.report_writer = ReportWriter()

    def run(
        self,
        input_dir: str | Path,
        dictionary_path: str | Path,
        output_dir: str | Path,
        mode: str = "markers",
    ) -> list[GenerationReportRow]:
        if mode not in {"markers", "blank"}:
            raise ValueError("El modo debe ser 'markers' o 'blank'.")

        source_dir = Path(input_dir)
        target_dir = Path(output_dir)
        target_dir.mkdir(parents=True, exist_ok=True)

        rules = self.dictionary_loader.load(dictionary_path, mode)
        rows: list[GenerationReportRow] = []
        for source in self._iter_documents(source_dir):
            destination = self._destination_path(source, source_dir, target_dir)
            destination.parent.mkdir(parents=True, exist_ok=True)
            rows.append(self._process_file(source, destination, rules))

        self.report_writer.write(target_dir, rows)
        return rows

    def _iter_documents(self, input_dir: Path) -> list[Path]:
        if not input_dir.is_dir():
            raise FileNotFoundError(f"No existe la carpeta de entrada: {input_dir}")
        return sorted(
            path
            for path in input_dir.rglob("*")
            if path.is_file()
            and path.suffix.lower() in self.SUPPORTED_SUFFIXES
            and not path.name.startswith("~$")
        )

    def _process_file(
        self,
        source: Path,
        destination: Path,
        rules: list[ReplacementRule],
    ) -> GenerationReportRow:
        suffix = source.suffix.lower()
        try:
            if suffix == ".docx":
                count, values = self.word_cleaner.clean(source, destination, rules)
            elif suffix in {".xlsx", ".xlsm"}:
                count, values = self.excel_cleaner.clean(source, destination, rules)
            elif suffix == ".pdf":
                count, values = self.pdf_cleaner.clean(source, destination, rules)
            else:
                raise ValueError(f"Tipo de archivo no soportado: {suffix}")
            return GenerationReportRow(
                archivo=str(source),
                tipo=suffix.lstrip("."),
                estado="ok",
                cantidad_reemplazos=count,
                valores_reemplazados="; ".join(values),
                ruta_salida=str(destination),
            )
        except Exception as exc:
            return GenerationReportRow(
                archivo=str(source),
                tipo=suffix.lstrip("."),
                estado="error",
                errores=str(exc),
                ruta_salida=str(destination),
            )

    def _destination_path(
        self,
        source: Path,
        input_dir: Path,
        output_dir: Path,
    ) -> Path:
        relative = source.relative_to(input_dir)
        destination = output_dir / relative
        if not destination.exists():
            return destination
        stem = destination.stem
        suffix = destination.suffix
        parent = destination.parent
        index = 2
        while True:
            candidate = parent / f"{stem}_{index}{suffix}"
            if not candidate.exists():
                return candidate
            index += 1
