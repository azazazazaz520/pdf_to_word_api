from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from docx import Document
from reportlab.pdfgen.canvas import Canvas

from src.validate.render import inspect_rendered_pdf
from src.validate.report import validate_docx_rendering


def _create_rendered_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Rendered page with text")
    canvas.showPage()
    canvas.showPage()
    canvas.save()


class DocxRenderValidationTest(unittest.TestCase):
    def test_inspect_rendered_pdf_reports_page_delta_and_blank_pages(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "rendered.pdf"
            _create_rendered_pdf(pdf_path)

            result = inspect_rendered_pdf(pdf_path, source_page_count=1)

            self.assertEqual(result["status"], "succeeded")
            self.assertEqual(result["rendered_page_count"], 2)
            self.assertEqual(result["page_delta"], 1)
            self.assertIn(2, result["blank_pages"])

    def test_validate_docx_rendering_uses_renderer_and_inspects_pdf(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            docx_path = root / "input.docx"
            Document().save(str(docx_path))
            rendered_source = root / "rendered-source.pdf"
            _create_rendered_pdf(rendered_source)

            def fake_render(
                _docx_path: Path,
                pdf_path: Path,
                *,
                timeout_seconds: float,
            ) -> None:
                pdf_path.write_bytes(rendered_source.read_bytes())

            with patch(
                "src.validate.render.render_docx_to_pdf",
                side_effect=fake_render,
            ):
                result = validate_docx_rendering(
                    docx_path,
                    source_page_count=1,
                    work_dir=root,
                )

            self.assertEqual(result["rendered_page_count"], 2)
            self.assertEqual(result["page_delta"], 1)
            self.assertEqual(result["blank_pages"], [2])
            self.assertTrue(Path(result["pdf_path"]).is_file())


if __name__ == "__main__":
    unittest.main()