from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app.plantilla_generator import TemplateGenerator


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera plantillas limpias desde documentos diligenciados.",
    )
    parser.add_argument("--input", default="data/entrada")
    parser.add_argument("--dictionary", default="diccionario_madecentro.xlsx")
    parser.add_argument("--output", default="plantillas_generadas")
    parser.add_argument(
        "--mode",
        choices=("markers", "blank"),
        default="markers",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = TemplateGenerator().run(
        input_dir=args.input,
        dictionary_path=args.dictionary,
        output_dir=args.output,
        mode=args.mode,
    )
    ok = sum(1 for row in rows if row.estado == "ok")
    errors = sum(1 for row in rows if row.estado == "error")
    print(f"Procesados: {len(rows)} | OK: {ok} | Errores: {errors}")
    print(f"Reporte: {Path(args.output) / 'reporte_generacion_plantillas.xlsx'}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
