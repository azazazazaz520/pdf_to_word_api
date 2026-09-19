from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from reportlab.pdfgen.canvas import Canvas

from src.worker.source_audit import audit_pdfium_source_characters


class SourceAuditTest(unittest.TestCase):
    def test_raw_pdfium_ids_can_be_checked_without_layout_objects(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "source-audit.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(300, 200))
            canvas.drawString(30, 150, "alpha beta")
            canvas.save()

            import pypdfium2 as pdfium

            document = pdfium.PdfDocument(str(pdf_path))
            try:
                text_page = document[0].get_textpage()
                try:
                    count = int(text_page.count_chars())
                finally:
                    text_page.close()
            finally:
                document.close()
            report = audit_pdfium_source_characters(
                pdf_path,
                source_char_ids=tuple(f"1:{index}" for index in range(count)),
            )

        self.assertTrue(report["independent_raw_pdfium_character_list"])
        self.assertEqual(report["status"], "passed")
        self.assertGreater(report["raw_visible_character_count"], 0)
        self.assertGreater(report["raw_space_count"], 0)

    def test_unknown_source_id_fails_independent_audit(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "source-audit.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(300, 200))
            canvas.drawString(30, 150, "alpha")
            canvas.save()

            report = audit_pdfium_source_characters(
                pdf_path,
                source_char_ids=("1:9999",),
            )

        self.assertEqual(report["status"], "failed")
        self.assertEqual(report["missing_source_character_count"], 1)
