from __future__ import annotations

from datetime import date
from hashlib import sha256
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree


WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class WordTemplateService:
    def compare_documents(
        self,
        empty_path: str | Path,
        completed_path: str | Path,
    ) -> dict[str, Any]:
        empty_path = Path(empty_path)
        completed_path = Path(completed_path)
        empty_root = self._read_document(empty_path)
        completed_root = self._read_document(completed_path)
        empty_slots = self._slots(empty_root)
        completed_slots = self._slots(completed_root)
        fields: list[dict[str, Any]] = []
        for field_id, empty_slot in empty_slots.items():
            completed_slot = completed_slots.get(field_id)
            if not completed_slot:
                continue
            if empty_slot["value"] == completed_slot["value"]:
                continue
            sample_value, value_format = self._extract_value(
                empty_slot["value"],
                completed_slot["value"],
            )
            label = empty_slot["label"]
            if (
                empty_slot["kind"] == "table_cell"
                and empty_slot["value"]
                and len(empty_slot["value"]) <= 160
            ):
                label = empty_slot["value"]
            field = {
                "field_id": field_id,
                "kind": empty_slot["kind"],
                "location": empty_slot["location"],
                "label": label,
                "empty_value": empty_slot["value"],
                "sample_value": sample_value,
                "master_key": "",
            }
            if value_format:
                field["value_format"] = value_format
            if self._is_checkbox_change(
                empty_slot["value"],
                completed_slot["value"],
            ):
                field["kind"] = "checkbox"
                field["preserve_reference"] = True
                field["reference_value"] = completed_slot["value"]
            fields.append(
                field
            )
        if not fields:
            raise ValueError(
                "No se detectaron campos estructurados distintos entre los DOCX."
            )
        payload = {
            "name": "mapeo_formulario_word",
            "format": "docx",
            "template_file": empty_path.name,
            "reference_file": completed_path.name,
            "fields": fields,
        }
        signature = self._detect_signature(empty_path, completed_path)
        if signature:
            payload["signature"] = signature
        return payload

    def fill_document(
        self,
        input_path: str | Path,
        output_path: str | Path,
        mapping: dict[str, Any],
        master_data: dict[str, str],
    ) -> Path:
        input_path = Path(input_path).resolve()
        output_path = Path(output_path).resolve()
        if input_path == output_path:
            raise ValueError("El Word de salida debe ser diferente al de entrada.")
        root = self._read_document(input_path)
        slots = self._slots(root)
        for field in mapping.get("fields", []):
            master_key = str(field.get("master_key") or "")
            slot = slots.get(str(field.get("field_id") or ""))
            if not slot:
                continue
            if field.get("value_transform") == "current_date":
                value = date.today().strftime(
                    str(field.get("date_format") or "%d/%m/%Y")
                )
            elif master_key:
                value = str(master_data.get(master_key, ""))
            elif field.get("preserve_reference"):
                value = str(
                    field.get("reference_value")
                    or field.get("sample_value")
                    or ""
                )
            else:
                continue
            value_format = str(field.get("value_format") or "")
            if value_format:
                value = value_format.replace("{value}", value)
            if field.get("kind") == "checkbox":
                self._set_checkbox(slot["element"], value)
            else:
                self._set_text(slot["element"], value)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = output_path.with_name(f".{output_path.stem}.tmp.docx")
        document_xml = etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )
        try:
            with ZipFile(input_path) as source, ZipFile(
                temp_path,
                "w",
                compression=ZIP_DEFLATED,
            ) as target:
                for item in source.infolist():
                    data = (
                        document_xml
                        if item.filename == "word/document.xml"
                        else source.read(item.filename)
                    )
                    target.writestr(self._copy_zip_info(item), data)
            temp_path.replace(output_path)
        finally:
            temp_path.unlink(missing_ok=True)
        return output_path

    def _read_document(self, path: Path) -> etree._Element:
        if not path.is_file() or path.suffix.lower() != ".docx":
            raise ValueError(f"El archivo no es un DOCX valido: {path}")
        with ZipFile(path) as archive:
            if "word/document.xml" not in archive.namelist():
                raise ValueError("El DOCX no contiene word/document.xml.")
            return etree.fromstring(archive.read("word/document.xml"))

    def _slots(self, root: etree._Element) -> dict[str, dict[str, Any]]:
        slots: dict[str, dict[str, Any]] = {}
        for index, control in enumerate(
            root.xpath(".//w:sdt", namespaces=WORD_NS),
            start=1,
        ):
            names = control.xpath(
                "./w:sdtPr/w:tag/@w:val | ./w:sdtPr/w:alias/@w:val",
                namespaces=WORD_NS,
            )
            label = str(names[0]) if names else f"Control {index}"
            field_id = f"sdt:{label}" if names else f"sdt-index:{index}"
            slots[field_id] = {
                "kind": "content_control",
                "location": field_id,
                "label": label,
                "value": self._text(control),
                "element": control,
            }

        for table_index, table in enumerate(
            root.xpath(".//w:tbl", namespaces=WORD_NS),
            start=1,
        ):
            rows = table.xpath("./w:tr", namespaces=WORD_NS)
            for row_index, row in enumerate(rows, start=1):
                cells = row.xpath("./w:tc", namespaces=WORD_NS)
                for column_index, cell in enumerate(cells, start=1):
                    if cell.xpath(".//w:sdt", namespaces=WORD_NS):
                        continue
                    field_id = f"table:{table_index}:{row_index}:{column_index}"
                    label = ""
                    if column_index > 1:
                        label = self._text(cells[column_index - 2])
                    if not label and row_index > 1:
                        previous_cells = rows[row_index - 2].xpath(
                            "./w:tc",
                            namespaces=WORD_NS,
                        )
                        if column_index <= len(previous_cells):
                            label = self._text(previous_cells[column_index - 1])
                    slots[field_id] = {
                        "kind": "table_cell",
                        "location": field_id,
                        "label": label or field_id,
                        "value": self._text(cell),
                        "element": cell,
                    }
        return slots

    def _text(self, element: etree._Element) -> str:
        return " ".join(
            value.strip()
            for value in element.xpath(".//w:t/text()", namespaces=WORD_NS)
            if value.strip()
        ).strip()

    def _set_text(self, element: etree._Element, value: str) -> None:
        text_nodes = element.xpath(".//w:t", namespaces=WORD_NS)
        if text_nodes:
            text_nodes[0].text = str(value)
            for node in text_nodes[1:]:
                node.text = ""
            return
        paragraph = element.find(".//w:p", namespaces=WORD_NS)
        if paragraph is None:
            paragraph = etree.SubElement(
                element,
                f"{{{WORD_NS['w']}}}p",
            )
        run = etree.SubElement(paragraph, f"{{{WORD_NS['w']}}}r")
        etree.SubElement(run, f"{{{WORD_NS['w']}}}t").text = str(value)

    def _set_checkbox(self, element: etree._Element, value: str) -> None:
        checked = str(value).strip() in {"☒", "☑", "1", "true", "True"}
        checked_nodes = element.xpath(
            "./w:sdtPr/w14:checkbox/w14:checked",
            namespaces=WORD_NS,
        )
        for node in checked_nodes:
            node.set(f"{{{WORD_NS['w14']}}}val", "1" if checked else "0")
        self._set_text(element, "☒" if checked else "☐")

    def _extract_value(self, empty: str, completed: str) -> tuple[str, str]:
        if empty and completed.startswith(empty):
            tail = completed[len(empty):]
            sample = tail.strip()
            if sample:
                separator = tail[: len(tail) - len(tail.lstrip())]
                return sample, f"{empty}{separator}{{value}}"
        return completed, ""

    def _is_checkbox_change(self, empty: str, completed: str) -> bool:
        marks = {"☐", "☒", "☑"}
        return empty in marks and completed in marks

    def _detect_signature(
        self,
        empty_path: Path,
        completed_path: Path,
    ) -> dict[str, Any] | None:
        with ZipFile(empty_path) as empty_archive:
            empty_hashes = {
                sha256(empty_archive.read(name)).digest()
                for name in empty_archive.namelist()
                if name.startswith("word/media/")
            }
        with ZipFile(completed_path) as archive:
            rels_name = "word/_rels/document.xml.rels"
            if rels_name not in archive.namelist():
                return None
            root = etree.fromstring(archive.read("word/document.xml"))
            relationships = etree.fromstring(archive.read(rels_name))
            targets = {
                str(relation.get("Id")): str(relation.get("Target"))
                for relation in relationships
                if str(relation.get("Type") or "").endswith("/image")
            }
            new_relations: set[str] = set()
            for relation_id, target in targets.items():
                media_name = f"word/{target.lstrip('/')}"
                if media_name not in archive.namelist():
                    continue
                digest = sha256(archive.read(media_name)).digest()
                if digest not in empty_hashes:
                    new_relations.add(relation_id)
            for blip in root.xpath(".//a:blip", namespaces=WORD_NS):
                relation_id = blip.get(f"{{{WORD_NS['r']}}}embed")
                if relation_id not in new_relations:
                    continue
                paragraph = self._ancestor(blip, "p")
                cell = self._ancestor(blip, "tc")
                row = self._ancestor(blip, "tr")
                table = self._ancestor(blip, "tbl")
                if paragraph is None or cell is None or row is None or table is None:
                    continue
                tables = root.xpath(".//w:tbl", namespaces=WORD_NS)
                rows = table.xpath("./w:tr", namespaces=WORD_NS)
                cells = row.xpath("./w:tc", namespaces=WORD_NS)
                paragraphs = cell.xpath("./w:p", namespaces=WORD_NS)
                extent = blip.xpath(
                    "ancestor::w:drawing[1]//wp:extent[1]",
                    namespaces={
                        **WORD_NS,
                        "wp": (
                            "http://schemas.openxmlformats.org/drawingml/"
                            "2006/wordprocessingDrawing"
                        ),
                    },
                )
                config: dict[str, Any] = {
                    "enabled": True,
                    "mode": "learned",
                    "table": tables.index(table) + 1,
                    "row": rows.index(row) + 1,
                    "cell": cells.index(cell) + 1,
                    "paragraph": paragraphs.index(paragraph) + 1,
                }
                if extent:
                    config["width_emu"] = int(extent[0].get("cx", "1900000"))
                    config["height_emu"] = int(extent[0].get("cy", "650000"))
                return config
        return None

    def _ancestor(
        self,
        element: etree._Element,
        local_name: str,
    ) -> etree._Element | None:
        expected = f"{{{WORD_NS['w']}}}{local_name}"
        current = element.getparent()
        while current is not None:
            if current.tag == expected:
                return current
            current = current.getparent()
        return None

    def _copy_zip_info(self, item: ZipInfo) -> ZipInfo:
        copied = ZipInfo(item.filename, item.date_time)
        copied.compress_type = item.compress_type
        copied.comment = item.comment
        copied.extra = item.extra
        copied.internal_attr = item.internal_attr
        copied.external_attr = item.external_attr
        copied.create_system = item.create_system
        return copied
