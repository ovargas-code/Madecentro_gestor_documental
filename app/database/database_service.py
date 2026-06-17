from __future__ import annotations

import sqlite3
import hashlib
from pathlib import Path
from typing import Any

from app.core.settings import DB_PATH, MAPPINGS_DIR, ensure_directories
from app.models.import_models import ImportedValue, ImportChange
from app.models.schemas import MasterData, TemplateRecord


class DatabaseService:
    def __init__(self, db_path: Path = DB_PATH) -> None:
        self.db_path = Path(db_path)
        ensure_directories()
        self.db_path.parent.mkdir(parents=True, exist_ok=True)

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def initialize(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS datos_maestros (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    clave TEXT NOT NULL UNIQUE,
                    valor TEXT NOT NULL DEFAULT '',
                    categoria TEXT NOT NULL DEFAULT 'general',
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS plantillas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE,
                    ruta_pdf TEXT NOT NULL,
                    descripcion TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS mapeos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER,
                    nombre TEXT NOT NULL,
                    ruta_json TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (plantilla_id) REFERENCES plantillas(id)
                );

                CREATE TABLE IF NOT EXISTS importaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    archivo TEXT NOT NULL,
                    formato TEXT NOT NULL,
                    cantidad INTEGER NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS importacion_detalles (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    importacion_id INTEGER NOT NULL,
                    clave TEXT NOT NULL,
                    valor_anterior TEXT NOT NULL DEFAULT '',
                    valor_nuevo TEXT NOT NULL DEFAULT '',
                    campo_origen TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY (importacion_id) REFERENCES importaciones(id)
                );

                CREATE TABLE IF NOT EXISTS plantilla_versiones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER NOT NULL,
                    version INTEGER NOT NULL,
                    ruta_plantilla TEXT NOT NULL,
                    ruta_referencia TEXT NOT NULL DEFAULT '',
                    ruta_manifiesto TEXT NOT NULL,
                    fingerprint TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (plantilla_id) REFERENCES plantillas(id),
                    UNIQUE (plantilla_id, version)
                );

                CREATE TABLE IF NOT EXISTS formularios_cargados (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    plantilla_id INTEGER,
                    archivo TEXT NOT NULL,
                    formato TEXT NOT NULL,
                    archivo_sha256 TEXT NOT NULL,
                    cantidad_respuestas INTEGER NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (plantilla_id) REFERENCES plantillas(id)
                );

                CREATE TABLE IF NOT EXISTS respuestas_formulario (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    formulario_id INTEGER NOT NULL,
                    clave TEXT NOT NULL,
                    valor TEXT NOT NULL DEFAULT '',
                    campo_origen TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY (formulario_id) REFERENCES formularios_cargados(id)
                );

                CREATE TABLE IF NOT EXISTS clientes_certificados (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    razon_social TEXT NOT NULL,
                    nit TEXT NOT NULL,
                    nit_normalizado TEXT NOT NULL UNIQUE,
                    anio_vinculacion TEXT NOT NULL,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """
            )
            self._migrate_template_formats(conn)
            self._migrate_legacy_mappings(conn)
            self._repair_mapping_paths(conn)

    def _migrate_template_formats(self, conn: sqlite3.Connection) -> None:
        columns = {
            str(row["name"])
            for row in conn.execute("PRAGMA table_info(plantillas)").fetchall()
        }
        if "formato" not in columns:
            conn.execute(
                "ALTER TABLE plantillas ADD COLUMN formato TEXT NOT NULL DEFAULT 'pdf'"
            )
        if "ruta_referencia" not in columns:
            conn.execute(
                "ALTER TABLE plantillas ADD COLUMN ruta_referencia TEXT NOT NULL DEFAULT ''"
            )
        conn.execute(
            """
            UPDATE plantillas
            SET formato = CASE
                WHEN lower(ruta_pdf) LIKE '%.xlsx' THEN 'xlsx'
                WHEN lower(ruta_pdf) LIKE '%.docx' THEN 'docx'
                ELSE 'pdf'
            END
            WHERE formato IS NULL OR formato = '' OR formato = 'pdf'
            """
        )

    def _migrate_legacy_mappings(self, conn: sqlite3.Connection) -> None:
        template_rows = conn.execute("SELECT id FROM plantillas ORDER BY id").fetchall()
        if len(template_rows) == 1:
            conn.execute(
                "UPDATE mapeos SET plantilla_id = ? WHERE plantilla_id IS NULL",
                (int(template_rows[0]["id"]),),
            )
        conn.execute(
            """
            DELETE FROM mapeos
            WHERE id NOT IN (
                SELECT MAX(id)
                FROM mapeos
                GROUP BY plantilla_id, nombre
            )
            """
        )

    def _repair_mapping_paths(self, conn: sqlite3.Connection) -> None:
        rows = conn.execute(
            "SELECT id, plantilla_id, nombre, ruta_json FROM mapeos"
        ).fetchall()
        for row in rows:
            current = Path(str(row["ruta_json"]))
            if current.is_file() or row["plantilla_id"] is None:
                continue
            template_id = int(row["plantilla_id"])
            normalized_name = "".join(
                char.casefold()
                for char in str(row["nombre"])
                if char.isalnum()
            )
            candidates = []
            for candidate in MAPPINGS_DIR.glob(f"plantilla_{template_id}_*.json"):
                normalized_candidate = "".join(
                    char.casefold()
                    for char in candidate.stem
                    if char.isalnum()
                )
                if normalized_name and normalized_name in normalized_candidate:
                    candidates.append(candidate)
            if len(candidates) == 1:
                conn.execute(
                    "UPDATE mapeos SET ruta_json = ? WHERE id = ?",
                    (str(candidates[0].resolve()), int(row["id"])),
                )
        conn.executescript(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_mapeos_plantilla_nombre
            ON mapeos(plantilla_id, nombre)
            WHERE plantilla_id IS NOT NULL;

            CREATE UNIQUE INDEX IF NOT EXISTS idx_mapeos_sin_plantilla_nombre
            ON mapeos(nombre)
            WHERE plantilla_id IS NULL;
            """
        )

    def upsert_master_data(self, item: MasterData) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO datos_maestros (clave, valor, categoria, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(clave) DO UPDATE SET
                    valor = excluded.valor,
                    categoria = excluded.categoria,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (item.clave, item.valor, item.categoria),
            )

    def update_master_data_by_id(self, item_id: int, item: MasterData) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE datos_maestros
                SET clave = ?, valor = ?, categoria = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (item.clave, item.valor, item.categoria, item_id),
            )

    def bulk_upsert_master_data(self, items: list[MasterData]) -> int:
        if not items:
            return 0
        with self._connect() as conn:
            conn.executemany(
                """
                INSERT INTO datos_maestros (clave, valor, categoria, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(clave) DO UPDATE SET
                    valor = excluded.valor,
                    categoria = excluded.categoria,
                    updated_at = CURRENT_TIMESTAMP
                """,
                [(item.clave, item.valor, item.categoria) for item in items],
            )
        return len(items)

    def apply_form_import(
        self,
        changes: list[ImportChange],
        source_path: str | Path,
    ) -> int:
        if not changes:
            return 0
        source = Path(source_path)
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO importaciones (archivo, formato, cantidad)
                VALUES (?, ?, ?)
                """,
                (str(source.resolve()), source.suffix.lower().lstrip("."), len(changes)),
            )
            import_id = int(cursor.lastrowid)
            conn.executemany(
                """
                INSERT INTO datos_maestros (clave, valor, categoria, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(clave) DO UPDATE SET
                    valor = excluded.valor,
                    categoria = excluded.categoria,
                    updated_at = CURRENT_TIMESTAMP
                """,
                [
                    (change.master_key, change.new_value, change.category)
                    for change in changes
                ],
            )
            conn.executemany(
                """
                INSERT INTO importacion_detalles (
                    importacion_id,
                    clave,
                    valor_anterior,
                    valor_nuevo,
                    campo_origen
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                [
                    (
                        import_id,
                        change.master_key,
                        change.current_value,
                        change.new_value,
                        change.source_field,
                    )
                    for change in changes
                ],
            )
        return len(changes)

    def record_form_submission(
        self,
        values: list[ImportedValue],
        source_path: str | Path,
        plantilla_id: int | None = None,
    ) -> int:
        source = Path(source_path)
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO formularios_cargados (
                    plantilla_id,
                    archivo,
                    formato,
                    archivo_sha256,
                    cantidad_respuestas
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    plantilla_id,
                    str(source.resolve()),
                    source.suffix.lower().lstrip("."),
                    digest,
                    len(values),
                ),
            )
            form_id = int(cursor.lastrowid)
            conn.executemany(
                """
                INSERT INTO respuestas_formulario (
                    formulario_id,
                    clave,
                    valor,
                    campo_origen
                )
                VALUES (?, ?, ?, ?)
                """,
                [
                    (form_id, item.master_key, item.value, item.source_field)
                    for item in values
                ],
            )
        return form_id

    def list_form_submissions(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, plantilla_id, archivo, formato, archivo_sha256,
                       cantidad_respuestas, created_at
                FROM formularios_cargados
                ORDER BY id DESC
                """
            ).fetchall()
        return [dict(row) for row in rows]

    def replace_certificate_customers(
        self,
        customers: list[dict[str, str]],
    ) -> int:
        with self._connect() as conn:
            conn.execute("DELETE FROM clientes_certificados")
            conn.executemany(
                """
                INSERT INTO clientes_certificados (
                    razon_social,
                    nit,
                    nit_normalizado,
                    anio_vinculacion,
                    updated_at
                )
                VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                [
                    (
                        item["cliente_razon_social"],
                        item["cliente_nit"],
                        item["nit_normalizado"],
                        item["cliente_anio_vinculacion"],
                    )
                    for item in customers
                ],
            )
        return len(customers)

    def upsert_certificate_customer(
        self,
        customer: dict[str, str],
    ) -> dict[str, Any]:
        normalized_nit = str(customer["nit_normalizado"])
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO clientes_certificados (
                    razon_social,
                    nit,
                    nit_normalizado,
                    anio_vinculacion,
                    updated_at
                )
                VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(nit_normalizado) DO UPDATE SET
                    razon_social = excluded.razon_social,
                    nit = excluded.nit,
                    anio_vinculacion = excluded.anio_vinculacion,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (
                    customer["cliente_razon_social"],
                    customer["cliente_nit"],
                    normalized_nit,
                    customer["cliente_anio_vinculacion"],
                ),
            )
            row = conn.execute(
                """
                SELECT id, razon_social, nit, nit_normalizado, anio_vinculacion
                FROM clientes_certificados
                WHERE nit_normalizado = ?
                """,
                (normalized_nit,),
            ).fetchone()
        return dict(row)

    def search_certificate_customers(
        self,
        query: str,
        limit: int = 30,
    ) -> list[dict[str, Any]]:
        text = query.strip()
        normalized_nit = "".join(char for char in text if char.isdigit())
        nit_pattern = f"%{normalized_nit}%" if normalized_nit else "__NO_NIT__"
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, razon_social, nit, nit_normalizado, anio_vinculacion
                FROM clientes_certificados
                WHERE lower(razon_social) LIKE lower(?)
                   OR nit_normalizado LIKE ?
                ORDER BY
                    CASE WHEN nit_normalizado = ? THEN 0 ELSE 1 END,
                    razon_social
                LIMIT ?
                """,
                (
                    f"%{text}%",
                    nit_pattern,
                    normalized_nit,
                    limit,
                ),
            ).fetchall()
        return [dict(row) for row in rows]

    def count_certificate_customers(self) -> int:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT COUNT(*) AS total FROM clientes_certificados"
            ).fetchone()
        return int(row["total"])

    def list_import_history(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, archivo, formato, cantidad, created_at
                FROM importaciones
                ORDER BY id DESC
                """
            ).fetchall()
        return [dict(row) for row in rows]

    def get_master_data(self) -> dict[str, str]:
        with self._connect() as conn:
            rows = conn.execute("SELECT clave, valor FROM datos_maestros ORDER BY clave").fetchall()
        return {row["clave"]: row["valor"] for row in rows}

    def list_master_data(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT id, clave, valor, categoria, updated_at FROM datos_maestros ORDER BY categoria, clave"
            ).fetchall()
        return [dict(row) for row in rows]

    def delete_master_data(self, clave: str) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM datos_maestros WHERE clave = ?", (clave,))

    def delete_master_data_by_id(self, item_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM datos_maestros WHERE id = ?", (item_id,))

    def delete_template(self, template_id: int) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                DELETE FROM respuestas_formulario
                WHERE formulario_id IN (
                    SELECT id FROM formularios_cargados WHERE plantilla_id = ?
                )
                """,
                (template_id,),
            )
            conn.execute("DELETE FROM formularios_cargados WHERE plantilla_id = ?", (template_id,))
            conn.execute("DELETE FROM plantilla_versiones WHERE plantilla_id = ?", (template_id,))
            conn.execute("DELETE FROM mapeos WHERE plantilla_id = ?", (template_id,))
            conn.execute("DELETE FROM plantillas WHERE id = ?", (template_id,))

    def add_template(self, template: TemplateRecord) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO plantillas (nombre, ruta_pdf, descripcion)
                VALUES (?, ?, ?)
                ON CONFLICT(nombre) DO UPDATE SET
                    ruta_pdf = excluded.ruta_pdf,
                    descripcion = excluded.descripcion,
                    formato = ?,
                    ruta_referencia = ?
                """,
                (
                    template.nombre,
                    template.ruta_pdf,
                    template.descripcion,
                    template.formato,
                    template.ruta_referencia,
                ),
            )
            conn.execute(
                """
                UPDATE plantillas
                SET formato = ?, ruta_referencia = ?
                WHERE nombre = ?
                """,
                (
                    template.formato,
                    template.ruta_referencia,
                    template.nombre,
                ),
            )
            if cur.lastrowid:
                return int(cur.lastrowid)
            row = conn.execute("SELECT id FROM plantillas WHERE nombre = ?", (template.nombre,)).fetchone()
            return int(row["id"])

    def list_templates(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, nombre, ruta_pdf, descripcion, formato,
                       ruta_referencia, created_at
                FROM plantillas
                ORDER BY id DESC
                """
            ).fetchall()
        return [dict(row) for row in rows]

    def add_template_version(
        self,
        plantilla_id: int,
        ruta_plantilla: str,
        ruta_referencia: str,
        ruta_manifiesto: str,
        fingerprint: str,
    ) -> int:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT COALESCE(MAX(version), 0) + 1 AS next_version
                FROM plantilla_versiones
                WHERE plantilla_id = ?
                """,
                (plantilla_id,),
            ).fetchone()
            version = int(row["next_version"])
            conn.execute(
                """
                INSERT INTO plantilla_versiones (
                    plantilla_id,
                    version,
                    ruta_plantilla,
                    ruta_referencia,
                    ruta_manifiesto,
                    fingerprint
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    plantilla_id,
                    version,
                    ruta_plantilla,
                    ruta_referencia,
                    ruta_manifiesto,
                    fingerprint,
                ),
            )
            return version

    def list_template_versions(self, plantilla_id: int) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, plantilla_id, version, ruta_plantilla,
                       ruta_referencia, ruta_manifiesto, fingerprint, created_at
                FROM plantilla_versiones
                WHERE plantilla_id = ?
                ORDER BY version DESC
                """,
                (plantilla_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def add_mapping_record(self, nombre: str, ruta_json: str, plantilla_id: int | None = None) -> int:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT id FROM mapeos
                WHERE nombre = ? AND (
                    plantilla_id = ? OR (plantilla_id IS NULL AND ? IS NULL)
                )
                ORDER BY id DESC
                LIMIT 1
                """,
                (nombre, plantilla_id, plantilla_id),
            ).fetchone()
            if row:
                mapping_id = int(row["id"])
                conn.execute(
                    """
                    UPDATE mapeos
                    SET ruta_json = ?, created_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (ruta_json, mapping_id),
                )
                return mapping_id
            cur = conn.execute(
                "INSERT INTO mapeos (plantilla_id, nombre, ruta_json) VALUES (?, ?, ?)",
                (plantilla_id, nombre, ruta_json),
            )
            return int(cur.lastrowid)

    def list_mappings(self, plantilla_id: int | None = None) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if plantilla_id is None:
                rows = conn.execute(
                    """
                    SELECT id, plantilla_id, nombre, ruta_json, created_at
                    FROM mapeos
                    ORDER BY id DESC
                    """
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT id, plantilla_id, nombre, ruta_json, created_at
                    FROM mapeos
                    WHERE plantilla_id = ?
                    ORDER BY id DESC
                    """,
                    (plantilla_id,),
                ).fetchall()
        return [dict(row) for row in rows]

