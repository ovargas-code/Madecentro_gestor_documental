from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app.database.database_service import DatabaseService
from app.services.bulk_template_import_service import BulkTemplateImportService
from app.services.form_template_service import FormTemplateService


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Importa masivamente plantillas desde una carpeta.",
    )
    parser.add_argument("carpeta", help="Carpeta con archivos vacios y diligenciados.")
    parser.add_argument(
        "--reporte",
        help="Ruta opcional para guardar el reporte CSV.",
    )
    parser.add_argument(
        "--use-openai",
        action="store_true",
        help="Usa OpenAI para sugerir mapeos de campos.",
    )
    args = parser.parse_args()

    load_dotenv(ROOT_DIR / ".env")
    if args.use_openai and not os.getenv("OPENAI_API_KEY", "").strip():
        parser.error(
            "--use-openai requiere OPENAI_API_KEY configurada en el entorno o en .env."
        )

    db = DatabaseService()
    db.initialize()
    service = BulkTemplateImportService(
        db=db,
        form_templates=FormTemplateService(use_openai_mapping=args.use_openai),
    )
    report = service.import_folder(args.carpeta, args.reporte)
    print(report.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
