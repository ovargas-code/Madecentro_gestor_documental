from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field, field_validator


class MasterData(BaseModel):
    clave: str = Field(min_length=1)
    valor: str = ""
    categoria: str = "general"

    @field_validator("clave", "categoria", mode="before")
    @classmethod
    def strip_text(cls, value: object) -> object:
        return value.strip() if isinstance(value, str) else value


class TemplateRecord(BaseModel):
    nombre: str = Field(min_length=1)
    ruta_pdf: str = Field(min_length=1)
    descripcion: str = ""
    formato: str = "pdf"
    ruta_referencia: str = ""

    @field_validator(
        "nombre",
        "ruta_pdf",
        "descripcion",
        "formato",
        "ruta_referencia",
        mode="before",
    )
    @classmethod
    def strip_text(cls, value: object) -> object:
        return value.strip() if isinstance(value, str) else value

    @field_validator("formato")
    @classmethod
    def validate_format(cls, value: str) -> str:
        normalized = value.lower().lstrip(".")
        if normalized not in {"pdf", "xlsx", "docx"}:
            raise ValueError("El formato debe ser PDF, XLSX o DOCX.")
        return normalized


class PdfField(BaseModel):
    field_name: str
    field_type: str
    page: int
    value: Optional[str] = None
    options: list[str] = Field(default_factory=list)


