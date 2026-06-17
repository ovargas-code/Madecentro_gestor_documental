import tempfile
import unittest
from pathlib import Path

import fitz

from app.services.pdf_field_service import PdfFieldService
from app.services.pdf_fill_service import (
    TEXT_BASE_SIZE,
    TEXT_MIN_SIZE,
    PdfFillService,
)
from app.services.pdf_signature_service import PdfSignatureService
from PIL import Image


def create_form(path: Path) -> None:
    document = fitz.open()
    page = document.new_page()

    text = fitz.Widget()
    text.field_name = "txt_nit"
    text.field_type = fitz.PDF_WIDGET_TYPE_TEXT
    text.rect = fitz.Rect(20, 20, 220, 50)
    text.text_font = "Cour"
    text.text_fontsize = 12
    text.text_color = (1, 0, 0)
    page.add_widget(text)

    long_text = fitz.Widget()
    long_text.field_name = "txt_largo"
    long_text.field_type = fitz.PDF_WIDGET_TYPE_TEXT
    long_text.rect = fitz.Rect(20, 55, 90, 75)
    page.add_widget(long_text)

    checkbox = fitz.Widget()
    checkbox.field_name = "chk_acepta"
    checkbox.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
    checkbox.rect = fitz.Rect(20, 90, 40, 110)
    page.add_widget(checkbox)

    document.save(path)
    document.close()


class PdfServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.input_path = Path(self.temp_dir.name) / "input.pdf"
        create_form(self.input_path)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_lists_and_fills_text_and_checkbox_fields(self) -> None:
        fields = PdfFieldService().list_fields(self.input_path)
        self.assertEqual(
            [(field.field_name, field.field_type) for field in fields],
            [
                ("txt_nit", "text"),
                ("txt_largo", "text"),
                ("chk_acepta", "checkbox"),
            ],
        )

        output_path = Path(self.temp_dir.name) / "output.pdf"
        PdfFillService().fill_pdf(
            self.input_path,
            output_path,
            {
                "txt_nit": "nit",
                "txt_largo": "descripcion",
                "chk_acepta": "acepta",
            },
            {
                "nit": "900123",
                "descripcion": "Texto suficientemente largo para ajustar",
                "acepta": "si",
            },
        )

        with fitz.open(output_path) as document:
            widgets = {
                widget.field_name: widget
                for page in document
                for widget in (page.widgets() or [])
            }
            acroform_type, acroform_value = document.xref_get_key(
                document.pdf_catalog(),
                "AcroForm",
            )
            if acroform_type == "xref":
                acroform_xref = int(acroform_value.split()[0])
                default_appearance = document.xref_get_key(acroform_xref, "DA")
                need_appearances = document.xref_get_key(
                    acroform_xref,
                    "NeedAppearances",
                )
            else:
                self.assertEqual(acroform_type, "dict")
                catalog_xref = document.pdf_catalog()
                default_appearance = document.xref_get_key(
                    catalog_xref,
                    "AcroForm/DA",
                )
                need_appearances = document.xref_get_key(
                    catalog_xref,
                    "AcroForm/NeedAppearances",
                )
            self.assertEqual(
                default_appearance,
                ("string", "/Helv 9 Tf 0 g"),
            )
            self.assertEqual(
                need_appearances,
                ("bool", "false"),
            )

            self.assertEqual(widgets["txt_nit"].field_value, "900123")
            self.assertEqual(widgets["txt_nit"].text_font, "Helv")
            self.assertEqual(widgets["txt_nit"].text_fontsize, TEXT_BASE_SIZE)
            self.assertEqual(widgets["txt_nit"].text_color, [0.0])
            self.assertEqual(
                document.xref_get_key(widgets["txt_nit"].xref, "DA"),
                ("string", "0 g /Helv 9 Tf"),
            )
            self.assertNotEqual(
                document.xref_get_key(widgets["txt_nit"].xref, "AP"),
                ("null", "null"),
            )

            self.assertGreaterEqual(
                widgets["txt_largo"].text_fontsize,
                TEXT_MIN_SIZE,
            )
            self.assertLess(
                widgets["txt_largo"].text_fontsize,
                TEXT_BASE_SIZE,
            )
            self.assertNotIn(
                widgets["chk_acepta"].field_value,
                ("", "Off", None),
            )

    def test_refuses_to_overwrite_input_pdf(self) -> None:
        with self.assertRaises(ValueError):
            PdfFillService().fill_pdf(self.input_path, self.input_path, {}, {})

    def test_adds_signature_image_to_pdf(self) -> None:
        signature = Path(self.temp_dir.name) / "signature.png"
        Image.new("RGB", (120, 40), "black").save(signature)
        output = Path(self.temp_dir.name) / "signed.pdf"
        PdfFillService().fill_pdf(self.input_path, output, {}, {})

        PdfSignatureService().apply_signature(output, signature)

        with fitz.open(output) as document:
            self.assertTrue(document[-1].get_images(full=True))

    def test_signature_replaces_acroform_placeholder(self) -> None:
        placeholder_pdf = Path(self.temp_dir.name) / "placeholder.pdf"
        document = fitz.open()
        page = document.new_page()
        widget = fitz.Widget()
        widget.field_name = "firma_png"
        widget.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
        widget.rect = fitz.Rect(20, 20, 180, 80)
        page.add_widget(widget)
        document.save(placeholder_pdf)
        document.close()
        signature = Path(self.temp_dir.name) / "signature-placeholder.png"
        Image.new("RGB", (120, 40), "black").save(signature)

        PdfSignatureService().apply_signature(placeholder_pdf, signature)

        with fitz.open(placeholder_pdf) as signed:
            field_names = {
                item.field_name
                for page in signed
                for item in (page.widgets() or [])
            }
            self.assertNotIn("firma_png", field_names)
            self.assertTrue(signed[0].get_images(full=True))


if __name__ == "__main__":
    unittest.main()
