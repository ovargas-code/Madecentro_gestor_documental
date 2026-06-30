from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook


@dataclass
class GenerationReportRow:
    archivo: str
    tipo: str
    estado: str
    cantidad_reemplazos: int = 0
    valores_reemplazados: str = ""
    errores: str = ""
    ruta_salida: str = ""


class ReportWriter:
    HEADERS = [
        "archivo",
        "tipo",
        "estado",
        "cantidad_reemplazos",
        "valores_reemplazados",
        "errores",
    ]

    def write(
        self,
        output_dir: str | Path,
        rows: list[GenerationReportRow],
    ) -> Path:
        output_path = Path(output_dir) / "reporte_generacion_plantillas.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Reporte"
        sheet.append(self.HEADERS)
        for row in rows:
            sheet.append(
                [
                    row.archivo,
                    row.tipo,
                    row.estado,
                    row.cantidad_reemplazos,
                    row.valores_reemplazados,
                    row.errores,
                ]
            )
        for column in sheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            sheet.column_dimensions[column[0].column_letter].width = min(
                max(max_length + 2, 12),
                80,
            )
        workbook.save(output_path)
        workbook.close()
        return output_path
