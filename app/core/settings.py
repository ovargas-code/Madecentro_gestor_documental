from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parents[2]
load_dotenv(BASE_DIR / ".env")

ASSETS_DIR = BASE_DIR / "assets"
LOGO_DIR = ASSETS_DIR / "logo"
SIGNATURE_DIR = ASSETS_DIR / "firma"
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "entrada"
OUTPUT_DIR = DATA_DIR / "salidas"
MASTER_DIR = DATA_DIR / "maestros"
MASTER_DATA_EXPORT_PATH = MASTER_DIR / "datos_maestros.csv"
CUSTOMER_CATALOG_PATH = MASTER_DIR / "clientes_certificados.xlsx"
LEARNING_DIR = DATA_DIR / "aprendizaje"
TEMPLATES_DIR = BASE_DIR / "plantillas"
PDF_TEMPLATES_DIR = TEMPLATES_DIR / "pdf"
EXCEL_TEMPLATES_DIR = TEMPLATES_DIR / "excel"
WORD_TEMPLATES_DIR = TEMPLATES_DIR / "word"
MAPPINGS_DIR = TEMPLATES_DIR / "mapeos"


def _path_from_env(name: str, default: Path) -> Path:
    raw_value = os.getenv(name, "").strip()
    if not raw_value:
        return default
    path = Path(raw_value).expanduser()
    return path if path.is_absolute() else BASE_DIR / path


DB_PATH = _path_from_env("DATABASE_PATH", MASTER_DIR / "madecentro.db")

REQUIRED_DIRECTORIES = [
    LOGO_DIR,
    SIGNATURE_DIR,
    INPUT_DIR,
    OUTPUT_DIR,
    MASTER_DIR,
    LEARNING_DIR,
    PDF_TEMPLATES_DIR,
    EXCEL_TEMPLATES_DIR,
    WORD_TEMPLATES_DIR,
    MAPPINGS_DIR,
]


def ensure_directories() -> None:
    for directory in REQUIRED_DIRECTORIES:
        directory.mkdir(parents=True, exist_ok=True)
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def first_image(directory: Path) -> Path | None:
    for pattern in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
        images = sorted(directory.glob(pattern))
        if images:
            return images[0]
    return None


def logo_path() -> Path | None:
    return first_image(LOGO_DIR)


def signature_path() -> Path | None:
    for pattern in ("00_firma_activa.png", "00_firma_activa.jpg", "00_firma_activa.jpeg", "00_firma_activa.webp"):
        active = SIGNATURE_DIR / pattern
        if active.is_file():
            return active
    return first_image(SIGNATURE_DIR)
