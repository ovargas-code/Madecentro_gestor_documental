from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from posixpath import normpath
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree as ET


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
VML_NS = "urn:schemas-microsoft-com:vml"
EXCEL_NS = "urn:schemas-microsoft-com:office:excel"
XML_NS = "http://www.w3.org/XML/1998/namespace"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
TRUE_VALUES = {"1", "true", "si", "sí", "yes", "x", "on", "checked", "marcado"}
SPANISH_MONTHS = (
    "",
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
)

class ExcelFillService:
    def fill_workbook(
        self,
        input_path: str | Path,
        output_path: str | Path,
        mapping: dict[str, Any],
        master_data: dict[str, str],
        use_sample_values: bool = False,
    ) -> Path:
        input_path = Path(input_path).resolve()
        output_path = Path(output_path).resolve()
        if not input_path.is_file():
            raise FileNotFoundError(f"No existe el Excel de entrada: {input_path}")
        if input_path.suffix.lower() != ".xlsx":
            raise ValueError(f"El archivo no es XLSX: {input_path}")
        if input_path == output_path:
            raise ValueError("El Excel de salida debe ser diferente al de entrada.")

        replacements = self._build_replacements(
            input_path,
            mapping,
            master_data,
            use_sample_values,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = output_path.with_name(f".{output_path.stem}.tmp.xlsx")
        try:
            self._write_package(input_path, temp_path, replacements)
            temp_path.replace(output_path)
        finally:
            temp_path.unlink(missing_ok=True)
        return output_path

    def _build_replacements(
        self,
        input_path: Path,
        mapping: dict[str, Any],
        master_data: dict[str, str],
        use_sample_values: bool,
    ) -> dict[str, bytes]:
        with ZipFile(input_path) as archive:
            sheet_paths = self._sheet_paths(archive)
            fields_by_sheet: dict[str, list[dict[str, Any]]] = {}
            for field in mapping.get("cells", []):
                fields_by_sheet.setdefault(str(field.get("sheet") or ""), []).append(
                    field
                )

            replacements: dict[str, bytes] = {}
            shared_strings = self._shared_strings(archive)
            sheet_names = set(sheet_paths)
            sheet_names.update(fields_by_sheet)
            for sheet_name in sheet_names:
                fields = fields_by_sheet.get(sheet_name, [])
                sheet_path = sheet_paths.get(sheet_name)
                if not sheet_path:
                    raise ValueError(
                        f"La plantilla Excel no contiene la hoja '{sheet_name}'."
                    )
                sheet_xml = ET.fromstring(archive.read(sheet_path))
                for field in fields:
                    value = self._mapped_value(
                        field,
                        master_data,
                        use_sample_values,
                    )
                    if value is None:
                        continue
                    self._set_cell_value(
                        sheet_xml,
                        str(field["cell"]),
                        value,
                        str(field.get("value_type", "string")),
                    )
                changed = bool(fields)
                if not use_sample_values:
                    changed = (
                        self._apply_inferred_current_dates(
                            sheet_xml,
                            shared_strings,
                        )
                        or changed
                    )
                if changed:
                    replacements[sheet_path] = ET.tostring(
                        sheet_xml,
                        encoding="utf-8",
                        xml_declaration=True,
                        standalone=True,
                    )

            controls_by_shape = {
                int(control["shape_id"]): control
                for control in mapping.get("controls", [])
            }
            vml_path = "xl/drawings/vmlDrawing1.vml"
            if controls_by_shape and vml_path in archive.namelist():
                vml_xml = ET.fromstring(archive.read(vml_path))
                vml_changed = False
                for shape in vml_xml.findall(f".//{{{VML_NS}}}shape"):
                    shape_id = int(shape.attrib["id"].rsplit("s", 1)[-1])
                    control = controls_by_shape.get(shape_id)
                    if not control:
                        continue
                    checked = self._control_value(
                        control,
                        master_data,
                        use_sample_values,
                    )
                    if checked is None:
                        continue
                    client_data = shape.find(f"{{{EXCEL_NS}}}ClientData")
                    if client_data is not None:
                        current_checked = client_data.find(
                            f"{{{EXCEL_NS}}}Checked"
                        ) is not None
                        if current_checked != checked:
                            self._set_vml_checked(client_data, checked)
                            vml_changed = True
                    prop_path = f"xl/ctrlProps/{control['control_id']}"
                    prop_xml = ET.fromstring(archive.read(prop_path))
                    prop_checked = prop_xml.attrib.get("checked") == "Checked"
                    if prop_checked != checked:
                        if checked:
                            prop_xml.attrib["checked"] = "Checked"
                        else:
                            prop_xml.attrib.pop("checked", None)
                        replacements[prop_path] = ET.tostring(
                            prop_xml,
                            encoding="utf-8",
                            xml_declaration=True,
                            standalone=True,
                        )

                if vml_changed:
                    replacements[vml_path] = ET.tostring(
                        vml_xml,
                        encoding="utf-8",
                        xml_declaration=True,
                        standalone=True,
                    )
            self._apply_drawing_marks(archive, mapping, replacements)
        return replacements

    def _apply_drawing_marks(
        self,
        archive: ZipFile,
        mapping: dict[str, Any],
        replacements: dict[str, bytes],
    ) -> None:
        marks_by_path: dict[str, list[dict[str, Any]]] = {}
        for mark in mapping.get("drawing_marks", []):
            path = str(mark.get("path") or "")
            xml = str(mark.get("xml") or "")
            if path and xml:
                marks_by_path.setdefault(path, []).append(mark)

        for drawing_path, marks in marks_by_path.items():
            if drawing_path not in archive.namelist():
                continue
            drawing = ET.fromstring(
                replacements.get(drawing_path, archive.read(drawing_path))
            )
            self._remove_existing_drawing_marks(drawing)
            next_id = self._next_drawing_id(drawing)
            for index, mark in enumerate(marks, start=1):
                anchor = ET.fromstring(str(mark["xml"]).encode("utf-8"))
                for element in anchor.findall(f".//{{{XDR_NS}}}cNvPr"):
                    element.attrib["id"] = str(next_id)
                    element.attrib["name"] = f"Madecentro Mark {index}"
                    next_id += 1
                drawing.append(anchor)
            replacements[drawing_path] = ET.tostring(
                drawing,
                encoding="utf-8",
                xml_declaration=True,
                standalone=True,
            )

    def _remove_existing_drawing_marks(self, drawing: ET.Element) -> None:
        for anchor in list(drawing):
            names = [
                str(element.attrib.get("name") or "")
                for element in anchor.findall(f".//{{{XDR_NS}}}cNvPr")
            ]
            if any(name.startswith("Madecentro Mark ") for name in names):
                drawing.remove(anchor)

    def _next_drawing_id(self, drawing: ET.Element) -> int:
        ids = [
            int(value)
            for element in drawing.findall(f".//{{{XDR_NS}}}cNvPr")
            if (value := str(element.attrib.get("id") or "")).isdigit()
        ]
        return max(ids, default=0) + 1

    def _sheet_paths(self, archive: ZipFile) -> dict[str, str]:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(
            archive.read("xl/_rels/workbook.xml.rels")
        )
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(
                f"{{{PACKAGE_REL_NS}}}Relationship"
            )
        }
        paths: dict[str, str] = {}
        for sheet in workbook.findall(f".//{{{SHEET_NS}}}sheet"):
            relation_id = sheet.attrib.get(
                f"{{{DOCUMENT_REL_NS}}}id",
                "",
            )
            target = targets.get(relation_id, "")
            if not target:
                continue
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            else:
                sheet_path = normpath(f"xl/{target}")
            paths[str(sheet.attrib.get("name", ""))] = sheet_path
        return paths

    def _mapped_value(
        self,
        field: dict[str, Any],
        master_data: dict[str, str],
        use_sample_values: bool,
    ) -> object | None:
        automatic_value = self._automatic_value(field)
        if automatic_value is not None:
            return automatic_value
        if use_sample_values:
            value: object = field.get("sample_value")
        else:
            master_key = str(field.get("master_key") or "")
            if not master_key:
                if self._preserves_reference(field):
                    return field.get("sample_value")
                return None
            value = master_data.get(master_key, "")
            if str(value or "").strip() == "" and "sample_value" in field:
                value = field.get("sample_value")
        value = self._transform_value(value, field.get("value_transform"))
        value_format = field.get("value_format")
        if value_format and not use_sample_values:
            return str(value_format).format(
                value=value,
                current_date=datetime.now().strftime("%d/%m/%Y"),
            )
        return value

    def _transform_value(
        self,
        value: object,
        transform: object,
    ) -> object:
        transform_name = str(transform or "")
        if transform_name not in {
            "tax_id_number",
            "tax_id_check_digit",
        }:
            return value
        text = str(value or "").strip()
        match = re.fullmatch(r"(.+?)[-\s]+([0-9])", text)
        if not match:
            return text if transform_name == "tax_id_number" else ""
        number, check_digit = match.groups()
        return (
            number.strip()
            if transform_name == "tax_id_number"
            else check_digit
        )

    def _preserves_reference(self, field: dict[str, Any]) -> bool:
        if field.get("preserve_reference"):
            return True
        sample = str(field.get("sample_value") or "").strip().casefold()
        empty = field.get("empty_value")
        return empty in (None, "") and sample in {"x", "✓"}

    def _control_value(
        self,
        control: dict[str, Any],
        master_data: dict[str, str],
        use_sample_values: bool,
    ) -> bool | None:
        master_key = str(control.get("master_key") or "")
        if master_key:
            return str(master_data.get(master_key, "")).strip().lower() in TRUE_VALUES
        if "sample_checked" in control:
            return bool(control.get("sample_checked"))
        return None

    def _automatic_value(self, field: dict[str, Any]) -> object | None:
        automatic = str(field.get("auto_value") or "")
        if not automatic:
            automatic = {
                "FORMULARIO!E9": "current_day",
                "FORMULARIO!G9": "current_month_name",
                "FORMULARIO!K9": "current_year",
            }.get(str(field.get("field_id") or ""), "")
        if not automatic:
            return None
        now = datetime.now()
        if automatic == "current_day":
            return now.day
        if automatic == "current_month_name":
            return SPANISH_MONTHS[now.month]
        if automatic == "current_month_number":
            return now.month
        if automatic == "current_year":
            return now.year
        if automatic == "current_date":
            return now.strftime("%d/%m/%Y")
        return None

    def _apply_inferred_current_dates(
        self,
        sheet_xml: ET.Element,
        shared_strings: list[str],
    ) -> bool:
        cell_texts = self._cell_texts(sheet_xml, shared_strings)
        changed = False
        for coordinate, text in cell_texts.items():
            column_letters, row_number = self._split_coordinate(coordinate)
            column_number = self._column_number(column_letters)
            if self._is_header_current_date_label(text, row_number):
                self._set_cell_value(
                    sheet_xml,
                    f"{self._column_letters(column_number + 1)}{row_number}",
                    datetime.now().strftime("%d/%m/%Y"),
                    "string",
                )
                changed = True
                continue
            if not self._is_current_date_label(text):
                continue
            changed = (
                self._apply_date_parts(
                    sheet_xml,
                    cell_texts,
                    row_number,
                    column_number,
                )
                or self._apply_single_date(
                    sheet_xml,
                    cell_texts,
                    row_number,
                    column_number,
                )
                or changed
            )
        return changed

    def _apply_date_parts(
        self,
        sheet_xml: ET.Element,
        cell_texts: dict[str, str],
        label_row: int,
        label_column: int,
    ) -> bool:
        now = datetime.now()
        values = {
            "dia": now.day,
            "mes": now.month,
            "ano": now.year,
        }
        changed = False
        for column in range(label_column + 1, label_column + 8):
            header = self._normalized_text(
                cell_texts.get(f"{self._column_letters(column)}{label_row}", "")
            )
            key = None
            if header == "dia":
                key = "dia"
            elif header == "mes":
                key = "mes"
            elif header in {"ano", "año"}:
                key = "ano"
            if key is None:
                continue
            self._set_cell_value(
                sheet_xml,
                f"{self._column_letters(column)}{label_row + 1}",
                values[key],
                "number",
            )
            changed = True
        return changed

    def _apply_single_date(
        self,
        sheet_xml: ET.Element,
        cell_texts: dict[str, str],
        label_row: int,
        label_column: int,
    ) -> bool:
        for row in (label_row, label_row + 1):
            for column in range(label_column + 1, label_column + 5):
                coordinate = f"{self._column_letters(column)}{row}"
                existing = str(cell_texts.get(coordinate) or "").strip()
                if existing and not self._looks_like_date(existing):
                    continue
                self._set_cell_value(
                    sheet_xml,
                    coordinate,
                    datetime.now().strftime("%d/%m/%Y"),
                    "string",
                )
                return True
        return False

    def _is_current_date_label(self, text: str) -> bool:
        normalized = self._normalized_text(text)
        if "fecha" not in normalized:
            return False
        return any(
            token in normalized
            for token in (
                "solicitud",
                "solicitur",
                "creacion",
                "diligenciamiento",
            )
        )

    def _is_header_current_date_label(self, text: str, row_number: int) -> bool:
        return row_number <= 10 and self._normalized_text(text) == "fecha"

    def _looks_like_date(self, text: str) -> bool:
        return bool(
            re.fullmatch(r"[0-9]{1,2}[/\\-][0-9]{1,2}[/\\-][0-9]{2,4}", text)
            or re.fullmatch(r"[0-9]{1,2}", text)
            or re.fullmatch(r"[0-9]{4}", text)
        )

    def _cell_texts(
        self,
        sheet_xml: ET.Element,
        shared_strings: list[str],
    ) -> dict[str, str]:
        texts: dict[str, str] = {}
        for cell in sheet_xml.findall(f".//{{{SHEET_NS}}}c"):
            coordinate = str(cell.attrib.get("r") or "")
            if not coordinate:
                continue
            text = self._cell_text(cell, shared_strings)
            if text:
                texts[coordinate] = text
        return texts

    def _cell_text(self, cell: ET.Element, shared_strings: list[str]) -> str:
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            return "".join(
                text_element.text or ""
                for text_element in cell.findall(f".//{{{SHEET_NS}}}t")
            )
        value = cell.find(f"{{{SHEET_NS}}}v")
        if value is None or value.text is None:
            return ""
        if cell_type == "s":
            try:
                return shared_strings[int(value.text)]
            except (IndexError, ValueError):
                return ""
        return value.text

    def _shared_strings(self, archive: ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return []
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings: list[str] = []
        for item in root.findall(f"{{{SHEET_NS}}}si"):
            strings.append(
                "".join(
                    text_element.text or ""
                    for text_element in item.findall(f".//{{{SHEET_NS}}}t")
                )
            )
        return strings

    def _normalized_text(self, text: str) -> str:
        normalized = unicodedata.normalize("NFKD", text.casefold())
        normalized = "".join(
            char for char in normalized if not unicodedata.combining(char)
        )
        return " ".join(re.sub(r"[^a-z0-9]+", " ", normalized).split())

    def _split_coordinate(self, coordinate: str) -> tuple[str, int]:
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", coordinate)
        if not match:
            raise ValueError(f"Coordenada Excel invalida: {coordinate}")
        letters, row = match.groups()
        return letters, int(row)

    def _column_letters(self, number: int) -> str:
        letters = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    def _set_cell_value(
        self,
        sheet_xml: ET.Element,
        coordinate: str,
        value: object,
        value_type: str,
    ) -> None:
        cell = self._find_or_create_cell(sheet_xml, coordinate)
        for child in list(cell):
            if child.tag in {
                f"{{{SHEET_NS}}}v",
                f"{{{SHEET_NS}}}f",
                f"{{{SHEET_NS}}}is",
            }:
                cell.remove(child)

        if value_type == "number":
            numeric_value = self._numeric_value(value)
            if numeric_value is not None:
                cell.attrib.pop("t", None)
                ET.SubElement(cell, f"{{{SHEET_NS}}}v").text = numeric_value
                return

        cell.attrib["t"] = "inlineStr"
        inline_string = ET.SubElement(cell, f"{{{SHEET_NS}}}is")
        text = ET.SubElement(inline_string, f"{{{SHEET_NS}}}t")
        string_value = "" if value is None else str(value)
        if string_value != string_value.strip():
            text.attrib[f"{{{XML_NS}}}space"] = "preserve"
        text.text = string_value

    def _numeric_value(self, value: object) -> str | None:
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return str(value)
        text = str(value or "").strip()
        if re.fullmatch(r"-?[0-9]+(?:\.[0-9]+)?", text):
            return text
        return None

    def _find_or_create_cell(
        self,
        sheet_xml: ET.Element,
        coordinate: str,
    ) -> ET.Element:
        cell = sheet_xml.find(
            f".//{{{SHEET_NS}}}c[@r='{coordinate}']"
        )
        if cell is not None:
            return cell

        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", coordinate)
        if not match:
            raise ValueError(f"Coordenada Excel invalida: {coordinate}")
        column_letters, row_text = match.groups()
        row_number = int(row_text)
        sheet_data = sheet_xml.find(f"{{{SHEET_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("La hoja no contiene sheetData.")

        row = sheet_data.find(f"{{{SHEET_NS}}}row[@r='{row_number}']")
        if row is None:
            row = ET.Element(f"{{{SHEET_NS}}}row", {"r": str(row_number)})
            rows = list(sheet_data)
            insert_at = next(
                (
                    index
                    for index, existing in enumerate(rows)
                    if int(existing.attrib.get("r", "0")) > row_number
                ),
                len(rows),
            )
            sheet_data.insert(insert_at, row)

        cell = ET.Element(f"{{{SHEET_NS}}}c", {"r": coordinate})
        column_number = self._column_number(column_letters)
        existing_cells = list(row)
        insert_at = next(
            (
                index
                for index, existing in enumerate(existing_cells)
                if self._column_number(
                    re.match(r"[A-Z]+", existing.attrib.get("r", "A")).group()
                )
                > column_number
            ),
            len(existing_cells),
        )
        row.insert(insert_at, cell)
        return cell

    def _column_number(self, letters: str) -> int:
        value = 0
        for letter in letters:
            value = value * 26 + ord(letter) - ord("A") + 1
        return value

    def _set_vml_checked(
        self,
        client_data: ET.Element,
        checked: bool,
    ) -> None:
        checked_element = client_data.find(f"{{{EXCEL_NS}}}Checked")
        if checked:
            if checked_element is None:
                checked_element = ET.Element(f"{{{EXCEL_NS}}}Checked")
                checked_element.text = "1"
                no_three_d = client_data.find(f"{{{EXCEL_NS}}}NoThreeD")
                index = (
                    list(client_data).index(no_three_d)
                    if no_three_d is not None
                    else len(client_data)
                )
                client_data.insert(index, checked_element)
        elif checked_element is not None:
            client_data.remove(checked_element)

    def _write_package(
        self,
        source: Path,
        target: Path,
        replacements: dict[str, bytes],
    ) -> None:
        with ZipFile(source, "r") as input_zip, ZipFile(
            target,
            "w",
            compression=ZIP_DEFLATED,
        ) as output_zip:
            for item in input_zip.infolist():
                data = replacements.get(item.filename, input_zip.read(item.filename))
                output_zip.writestr(self._copy_zip_info(item), data)

    def _copy_zip_info(self, item: ZipInfo) -> ZipInfo:
        copied = ZipInfo(item.filename, item.date_time)
        copied.compress_type = item.compress_type
        copied.comment = item.comment
        copied.extra = item.extra
        copied.internal_attr = item.internal_attr
        copied.external_attr = item.external_attr
        copied.create_system = item.create_system
        return copied
