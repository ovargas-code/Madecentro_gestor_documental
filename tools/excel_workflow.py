from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

from app.database.database_service import DatabaseService
from app.services.excel_fill_service import ExcelFillService
from app.services.excel_template_service import ExcelTemplateService


EMPTY_PATH = PROJECT_DIR / "data" / "entrada" / "Sarlaft-Somer-Incare vacio.xlsx"
COMPLETED_PATH = (
    PROJECT_DIR / "data" / "entrada" / "Sarlaft-Somer-Incare diligenciado.xlsx"
)
MAPPING_PATH = (
    PROJECT_DIR / "plantillas" / "mapeos" / "mapeo_formulario_excel.json"
)
VERIFICATION_PATH = (
    PROJECT_DIR / "data" / "salidas" / "formulario_excel_verificacion.xlsx"
)
OUTPUT_PATH = (
    PROJECT_DIR
    / "data"
    / "salidas"
    / "formulario_excel_diligenciado_madecentro.xlsx"
)


def learn() -> None:
    mapping = ExcelTemplateService().compare_workbooks(
        EMPTY_PATH,
        COMPLETED_PATH,
    )
    MAPPING_PATH.write_text(
        json.dumps(mapping, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    ExcelFillService().fill_workbook(
        EMPTY_PATH,
        VERIFICATION_PATH,
        mapping,
        {},
        use_sample_values=True,
    )
    print(f"Mapeo: {MAPPING_PATH}")
    print(f"Verificacion: {VERIFICATION_PATH}")
    print(
        f"Campos: {len(mapping['cells'])}; "
        f"controles: {len(mapping['controls'])}"
    )


def fill() -> None:
    mapping = json.loads(MAPPING_PATH.read_text(encoding="utf-8"))
    ExcelFillService().fill_workbook(
        EMPTY_PATH,
        OUTPUT_PATH,
        mapping,
        DatabaseService().get_master_data(),
    )
    print(f"Resultado: {OUTPUT_PATH}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Aprende y diligencia el formulario Excel."
    )
    parser.add_argument("action", choices=("learn", "fill"))
    args = parser.parse_args()
    if args.action == "learn":
        learn()
    else:
        fill()


if __name__ == "__main__":
    main()
