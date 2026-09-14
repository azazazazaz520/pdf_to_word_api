from __future__ import annotations

import json
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from src.pdf_worker import process_job


def _base_payload(root: Path, pdf_path: Path, page_count: int) -> dict:
    return {
        "job_id": "fidelity-test",
        "filename": pdf_path.name,
        "input_path": str(pdf_path.resolve()),
        "output_path": str((root / "result.docx").resolve()),
        "progress_path": str((root / "progress.json").resolve()),
        "stage_log_path": str((root / "stages.jsonl").resolve()),
        "cancel_path": str((root / "cancel.requested").resolve()),
        "page_count": page_count,
        "engine": "structure-lite",
        "route_mode": "auto",
        "text_min_page_chars": 20,
        "text_min_page_ratio": 0.6,
        "text_high_quality_ratio": 0.8,
        "text_full_page_image_min_pixels": 300000,
        "text_garbled_char_ratio": 0.05,
        "page_image_max_pixels": 4 * 1024 * 1024,
        "page_image_jpeg_quality": 88,
        "task_timeout_seconds": 300.0,
        "ocr_time_budget_seconds": 60.0,
        "render_validation": False,
        "quality_gate_enabled": False,
    }


def _write_text_pdf(path: Path, page_count: int = 1) -> None:
    canvas = Canvas(str(path), pagesize=(595.0, 842.0))
    for page_number in range(1, page_count + 1):
        canvas.setFont("Helvetica", 11)
        canvas.drawString(
            72,
            760,
            f"Editable text page {page_number} with enough characters.",
        )
        canvas.drawString(72, 730, "Second line of the same page.")
        canvas.showPage()
    canvas.save()


def _write_dense_text_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(595.0, 842.0))
    canvas.setFont("Helvetica", 9)
    for line_number in range(60):
        canvas.drawString(
            72,
            800 - line_number * 12,
            f"Dense text line {line_number} remains editable in the output.",
        )
    canvas.save()


def _write_mixed_pdf(path: Path, image_path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "First page is editable text with enough characters.")
    canvas.showPage()
    canvas.drawImage(
        ImageReader(str(image_path)), 0, 0, width=360, height=240
    )
    canvas.drawString(36, 200, "Small overlay text on a full page image")
    canvas.showPage()
    canvas.save()


class PdfWorkerFidelityTest(unittest.TestCase):
    def test_conversion_defaults_to_structured_flow(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "text.pdf"
            _write_text_pdf(pdf_path, page_count=1)
            payload = _base_payload(root, pdf_path, page_count=1)

            result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["export_mode"], "structured")
            self.assertIsNotNone(quality.get("structured"))
            self.assertGreater(quality["structured"]["text_paragraph_count"], 0)
            with zipfile.ZipFile(root / "result.docx") as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertNotIn("<w:framePr", document_xml)

    def test_explicit_fidelity_mode_remains_available(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "text.pdf"
            _write_text_pdf(pdf_path, page_count=1)
            payload = _base_payload(root, pdf_path, page_count=1)
            payload["export_mode"] = "fidelity"

            result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["export_mode"], "fidelity")
            self.assertIsNotNone(quality.get("fidelity"))
            with zipfile.ZipFile(root / "result.docx") as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("<w:framePr", document_xml)

    def test_dense_text_page_stays_on_text_route(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "dense-text.pdf"
            _write_dense_text_pdf(pdf_path)
            payload = _base_payload(root, pdf_path, page_count=1)

            result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["route_summary"], {"text": 1})
            self.assertEqual(quality["visual_only_page_count"], 0)
            self.assertEqual(quality["page_results"][0]["route"], "text")
            self.assertTrue(quality["page_results"][0]["editable"])

    def test_fidelity_keeps_text_for_visual_pages(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "full-page.png"
            Image.new("RGB", (720, 480), color="white").save(image_path)
            pdf_path = root / "mixed.pdf"
            _write_mixed_pdf(pdf_path, image_path)
            payload = _base_payload(root, pdf_path, page_count=2)
            payload["export_mode"] = "fidelity"

            result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["route_summary"], {"text": 2})
            self.assertEqual(quality["visual_only_page_count"], 0)
            fidelity = quality["fidelity"]
            self.assertEqual(fidelity["page_image_fallback_pages"], [])
            self.assertEqual(fidelity["fallback_region_count"], 0)
            page_two = next(
                item for item in fidelity["pages"] if item["page"] == 2
            )
            self.assertEqual(page_two["fallback_regions"], [])
            self.assertTrue(page_two["editable"])
            with zipfile.ZipFile(root / "result.docx") as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("<wp:anchor", document_xml)

    def test_render_validation_never_replaces_text_with_page_image(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "low-coverage.pdf"
            _write_text_pdf(pdf_path, page_count=1)
            payload = _base_payload(root, pdf_path, page_count=1)
            payload["export_mode"] = "fidelity"
            payload["render_validation"] = True
            payload["fidelity_auto_fallback"] = True

            render_result = {
                "status": "succeeded",
                "source_page_count": 1,
                "rendered_page_count": 1,
                "page_delta": 0,
                "blank_pages": [],
                "unexpected_blank_pages": [],
                "text_coverage": {
                    "status": "succeeded",
                    "pages": [{"page": 1, "coverage": 0.1}],
                },
            }
            with patch.object(
                __import__("src.pdf_worker", fromlist=["validate_docx_rendering"]),
                "validate_docx_rendering",
                return_value=render_result,
            ) as validate:
                result = process_job(payload)

            self.assertEqual(validate.call_count, 1)
            quality = result["quality"]
            self.assertEqual(quality["fidelity_auto_fallback_pages"], [])
            self.assertEqual(quality["fidelity"]["page_image_fallback_pages"], [])
            self.assertEqual(quality["visual_only_page_count"], 0)

if __name__ == "__main__":
    unittest.main()
