from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
IMAGE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


class WordSignatureService:
    def apply_signature(
        self,
        document_path: str | Path,
        signature_path: str | Path,
        config: dict[str, Any] | None = None,
    ) -> Path:
        document_path = Path(document_path).resolve()
        signature_path = Path(signature_path).resolve()
        config = config or {}
        temp_path = document_path.with_name(
            f".{document_path.stem}.signature.tmp.docx"
        )
        try:
            with ZipFile(document_path) as source:
                root = etree.fromstring(source.read("word/document.xml"))
                rels_path = "word/_rels/document.xml.rels"
                relationships = (
                    etree.fromstring(source.read(rels_path))
                    if rels_path in source.namelist()
                    else etree.Element(f"{{{PACKAGE_REL_NS}}}Relationships")
                )
                relation_id = self._add_relationship(relationships)
                paragraph = self._signature_paragraph(root, config)
                paragraph.append(self._drawing_run(relation_id, config))
                replacements = {
                    "word/document.xml": etree.tostring(
                        root,
                        encoding="UTF-8",
                        xml_declaration=True,
                        standalone=True,
                    ),
                    rels_path: etree.tostring(
                        relationships,
                        encoding="UTF-8",
                        xml_declaration=True,
                        standalone=True,
                    ),
                    "word/media/madecentro_signature.png": signature_path.read_bytes(),
                }
                if "[Content_Types].xml" in source.namelist():
                    content_types = etree.fromstring(
                        source.read("[Content_Types].xml")
                    )
                    self._ensure_png_content_type(content_types)
                    replacements["[Content_Types].xml"] = etree.tostring(
                        content_types,
                        encoding="UTF-8",
                        xml_declaration=True,
                        standalone=True,
                    )
                self._write_package(source, temp_path, replacements)
            temp_path.replace(document_path)
        finally:
            temp_path.unlink(missing_ok=True)
        return document_path

    def _signature_paragraph(
        self,
        root: etree._Element,
        config: dict[str, Any],
    ) -> etree._Element:
        table_number = int(config.get("table", 0) or 0)
        row_number = int(config.get("row", 0) or 0)
        cell_number = int(config.get("cell", 0) or 0)
        paragraph_number = int(config.get("paragraph", 0) or 0)
        if table_number and row_number and cell_number:
            tables = root.xpath(".//w:tbl", namespaces={"w": WORD_NS})
            if table_number <= len(tables):
                rows = tables[table_number - 1].xpath(
                    "./w:tr",
                    namespaces={"w": WORD_NS},
                )
                if row_number <= len(rows):
                    cells = rows[row_number - 1].xpath(
                        "./w:tc",
                        namespaces={"w": WORD_NS},
                    )
                    if cell_number <= len(cells):
                        paragraphs = cells[cell_number - 1].xpath(
                            "./w:p",
                            namespaces={"w": WORD_NS},
                        )
                        if paragraph_number and paragraph_number <= len(paragraphs):
                            return paragraphs[paragraph_number - 1]
                        if paragraphs:
                            return paragraphs[-1]
        marker = str(config.get("marker") or "firma").casefold()
        for text in root.xpath(".//w:t", namespaces={"w": WORD_NS}):
            if marker in str(text.text or "").casefold():
                paragraph = text
                while paragraph is not None and paragraph.tag != f"{{{WORD_NS}}}p":
                    paragraph = paragraph.getparent()
                if paragraph is not None:
                    return paragraph
        body = root.find(f"{{{WORD_NS}}}body")
        if body is None:
            raise ValueError("El DOCX no contiene un cuerpo de documento.")
        paragraph = etree.Element(f"{{{WORD_NS}}}p")
        section = body.find(f"{{{WORD_NS}}}sectPr")
        if section is None:
            body.append(paragraph)
        else:
            body.insert(list(body).index(section), paragraph)
        return paragraph

    def _add_relationship(self, relationships: etree._Element) -> str:
        used = {
            int(value[3:])
            for relation in relationships
            if (value := str(relation.get("Id") or "")).startswith("rId")
            and value[3:].isdigit()
        }
        relation_id = f"rId{max(used, default=0) + 1}"
        etree.SubElement(
            relationships,
            f"{{{PACKAGE_REL_NS}}}Relationship",
            {
                "Id": relation_id,
                "Type": IMAGE_REL_TYPE,
                "Target": "media/madecentro_signature.png",
            },
        )
        return relation_id

    def _ensure_png_content_type(self, content_types: etree._Element) -> None:
        for item in content_types:
            if (
                item.tag == f"{{{CONTENT_TYPES_NS}}}Default"
                and str(item.get("Extension") or "").casefold() == "png"
            ):
                return
        etree.SubElement(
            content_types,
            f"{{{CONTENT_TYPES_NS}}}Default",
            {"Extension": "png", "ContentType": "image/png"},
        )

    def _drawing_run(
        self,
        relation_id: str,
        config: dict[str, Any],
    ) -> etree._Element:
        width = int(config.get("width_emu", 1900000))
        height = int(config.get("height_emu", 650000))
        run = etree.Element(f"{{{WORD_NS}}}r")
        drawing = etree.SubElement(run, f"{{{WORD_NS}}}drawing")
        inline = etree.SubElement(
            drawing,
            f"{{{WP_NS}}}inline",
            distT="0",
            distB="0",
            distL="0",
            distR="0",
        )
        etree.SubElement(
            inline,
            f"{{{WP_NS}}}extent",
            cx=str(width),
            cy=str(height),
        )
        etree.SubElement(
            inline,
            f"{{{WP_NS}}}docPr",
            id="50000",
            name="Firma Madecentro",
        )
        graphic = etree.SubElement(inline, f"{{{A_NS}}}graphic")
        graphic_data = etree.SubElement(
            graphic,
            f"{{{A_NS}}}graphicData",
            uri="http://schemas.openxmlformats.org/drawingml/2006/picture",
        )
        picture = etree.SubElement(graphic_data, f"{{{PIC_NS}}}pic")
        nv = etree.SubElement(picture, f"{{{PIC_NS}}}nvPicPr")
        etree.SubElement(
            nv,
            f"{{{PIC_NS}}}cNvPr",
            id="0",
            name="Firma Madecentro",
        )
        etree.SubElement(nv, f"{{{PIC_NS}}}cNvPicPr")
        fill = etree.SubElement(picture, f"{{{PIC_NS}}}blipFill")
        etree.SubElement(
            fill,
            f"{{{A_NS}}}blip",
            {f"{{{REL_NS}}}embed": relation_id},
        )
        stretch = etree.SubElement(fill, f"{{{A_NS}}}stretch")
        etree.SubElement(stretch, f"{{{A_NS}}}fillRect")
        shape = etree.SubElement(picture, f"{{{PIC_NS}}}spPr")
        transform = etree.SubElement(shape, f"{{{A_NS}}}xfrm")
        etree.SubElement(transform, f"{{{A_NS}}}off", x="0", y="0")
        etree.SubElement(
            transform,
            f"{{{A_NS}}}ext",
            cx=str(width),
            cy=str(height),
        )
        geometry = etree.SubElement(
            shape,
            f"{{{A_NS}}}prstGeom",
            prst="rect",
        )
        etree.SubElement(geometry, f"{{{A_NS}}}avLst")
        return run

    def _write_package(
        self,
        source: ZipFile,
        target_path: Path,
        replacements: dict[str, bytes],
    ) -> None:
        written: set[str] = set()
        with ZipFile(target_path, "w", compression=ZIP_DEFLATED) as target:
            for item in source.infolist():
                target.writestr(
                    self._copy_zip_info(item),
                    replacements.get(item.filename, source.read(item.filename)),
                )
                written.add(item.filename)
            for name, data in replacements.items():
                if name not in written:
                    target.writestr(name, data)

    def _copy_zip_info(self, item: ZipInfo) -> ZipInfo:
        copied = ZipInfo(item.filename, item.date_time)
        copied.compress_type = item.compress_type
        copied.comment = item.comment
        copied.extra = item.extra
        copied.internal_attr = item.internal_attr
        copied.external_attr = item.external_attr
        copied.create_system = item.create_system
        return copied
