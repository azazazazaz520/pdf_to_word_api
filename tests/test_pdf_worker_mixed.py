from __future__ import annotations

import json
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from unittest.mock import patch

import src.pdf_worker as pdf_worker
from src.pdf_worker import _evaluate_quality_gate, process_job


class _FakeOcrPipeline:
    def __init__(self) -> None:
        self.calls: list[str] = []

    def predict_iter(self, path: str):
        self.calls.append(path)
        yield {
            "overall_ocr_res": {
                "rec_texts": ["OCR page text"],
                "rec_scores": [0.98],
            },
            "parsing_res_list": [
                {
                    "block_label": "text",
                    "block_content": "OCR second page text",
                }
            ],
        }


def _base_payload(root: Path, pdf_path: Path, page_count: int) -> dict:
    return {
        "job_id": "mixed-test",
        "filename": pdf_path.name,
        "input_path": str(pdf_path),
        "output_path": str(root / "result.docx"),
        "progress_path": str(root / "progress.json"),
        "stage_log_path": str(root / "stages.jsonl"),
        "cancel_path": str(root / "cancel.requested"),
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
    }


class PdfWorkerMixedRouteTest(unittest.TestCase):
    def test_text_page_keeps_embedded_image(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "embedded.png"
            pdf_path = root / "text-with-image.pdf"
            Image.new("RGB", (200, 100), color=(20, 80, 160)).save(image_path)

            canvas = Canvas(str(pdf_path), pagesize=(400, 300))
            canvas.drawImage(
                ImageReader(str(image_path)),
                50,
                50,
                width=200,
                height=100,
            )
            canvas.drawString(50, 270, "Editable text content " * 8)
            canvas.save()

            result = process_job(_base_payload(root, pdf_path, page_count=1))

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["route_summary"], {"text": 1})
            self.assertEqual(quality["media_count"], 1)
            with zipfile.ZipFile(root / "result.docx") as archive:
                media = [
                    name
                    for name in archive.namelist()
                    if name.startswith("word/media/")
                ]
            self.assertEqual(len(media), 1)

    def test_mixed_document_uses_text_and_page_image_routes(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "full-page.png"
            pdf_path = root / "mixed.pdf"
            Image.new("RGB", (720, 480), color="white").save(image_path)

            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.drawString(
                36,
                200,
                "First page is editable text with enough characters.",
            )
            canvas.showPage()
            canvas.drawImage(
                ImageReader(str(image_path)),
                0,
                0,
                width=360,
                height=240,
            )
            canvas.drawString(36, 200, "Small overlay text on a full page image")
            canvas.save()

            result = process_job(_base_payload(root, pdf_path, page_count=2))

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(
                quality["route_summary"],
                {"text": 1, "page_image": 1},
            )
            self.assertEqual(quality["media_count"], 1)
            self.assertEqual(quality["needs_review_pages"], [2])
            page_routes = [item["route"] for item in quality["page_results"]]
            self.assertEqual(page_routes, ["text", "page_image"])

            stage_log = (root / "stages.jsonl").read_text(encoding="utf-8")
            events = [json.loads(line) for line in stage_log.splitlines()]
            self.assertTrue(
                any(event.get("stage") == "page_image_completed" for event in events)
            )


    def test_mixed_document_can_combine_text_and_ocr_pages(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "text-and-blank.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.drawString(
                36,
                200,
                "First page has selectable text with enough characters.",
            )
            canvas.showPage()
            canvas.showPage()
            canvas.save()

            payload = _base_payload(root, pdf_path, page_count=2)
            fake_pipeline = _FakeOcrPipeline()
            with patch.object(
                pdf_worker,
                "_get_pipeline",
                return_value=fake_pipeline,
            ):
                result = process_job(payload)
            self.assertEqual(len(fake_pipeline.calls), 1)
            self.assertTrue(
                fake_pipeline.calls[0].lower().endswith((".jpg", ".jpeg", ".png"))
            )

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["route_summary"], {"text": 1, "ocr": 1})
            self.assertEqual(
                [item["route"] for item in quality["page_results"]],
                ["text", "ocr"],
            )
            paragraphs = "\n".join(
                paragraph.text
                for paragraph in __import__("docx").Document(
                    str(root / "result.docx")
                ).paragraphs
            )
            self.assertIn("First page has selectable text", paragraphs)
            fidelity = quality["fidelity"]
            self.assertIn(2, fidelity["page_image_fallback_pages"])
            ocr_page = next(
                item for item in fidelity["pages"] if item["page"] == 2
            )
            self.assertIn(
                "ocr_page_preserved_as_image",
                [region["reason"] for region in ocr_page["fallback_regions"]],
            )



    def test_worker_attaches_render_validation_quality(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "render-validation.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.drawString(36, 200, "Render validation text page.")
            canvas.save()
            payload = _base_payload(root, pdf_path, page_count=1)
            payload["render_validation"] = True
            payload["render_timeout_seconds"] = 60

            with patch.object(
                pdf_worker,
                "validate_docx_rendering",
                return_value={
                    "status": "succeeded",
                    "source_page_count": 1,
                    "rendered_page_count": 2,
                    "page_delta": 1,
                    "blank_pages": [2],
                    "pdf_path": str(root / "result.rendered.pdf"),
                },
            ):
                result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            quality = result["quality"]
            self.assertEqual(quality["page_delta"], 1)
            self.assertEqual(quality["blank_pages"], [2])
            self.assertEqual(
                quality["render_validation"]["status"],
                "succeeded",
            )
            warning_codes = {
                warning["code"] for warning in quality["warnings"]
            }
            self.assertIn("render_page_mismatch", warning_codes)
            self.assertIn("render_blank_page", warning_codes)
            self.assertIn(1, quality["needs_review_pages"])
            self.assertIn(2, quality["needs_review_pages"])



    def test_worker_writes_pdf_outline_as_word_bookmarks_and_toc(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "outline.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.bookmarkPage("chapter1")
            canvas.addOutlineEntry("Chapter 1", "chapter1", level=0)
            canvas.drawString(36, 200, "Chapter one text with enough characters.")
            canvas.save()
            payload = _base_payload(root, pdf_path, page_count=1)
            payload["include_bookmarks"] = True
            payload["include_toc"] = True

            result = process_job(payload)

            self.assertEqual(result["status"], "succeeded")
            with zipfile.ZipFile(root / "result.docx") as archive:
                document_xml = archive.read("word/document.xml")
            self.assertIn(b"w:bookmarkStart", document_xml)
            self.assertIn(b"Chapter_1", document_xml)
            self.assertIn(b"TOC", document_xml)



    def test_quality_gate_flags_page_delta_and_blank_pages(self) -> None:
        quality = {
            "render_validation": {
                "status": "succeeded",
                "source_page_count": 100,
                "page_delta": 10,
                "blank_pages": [],
            }
        }
        gate = _evaluate_quality_gate(
            quality,
            enabled=True,
            page_delta_warn_ratio=0.05,
            page_delta_warn_absolute=3,
        )
        self.assertEqual(gate["status"], "failed")
        self.assertEqual(gate["page_delta_threshold"], 5)

        quality["render_validation"]["page_delta"] = 2
        quality["render_validation"]["blank_pages"] = [4]
        gate = _evaluate_quality_gate(
            quality,
            enabled=True,
            page_delta_warn_ratio=0.05,
            page_delta_warn_absolute=3,
        )
        self.assertEqual(gate["status"], "failed")
        self.assertTrue(
            any(
                check["name"] == "blank_pages"
                and check["status"] == "failed"
                for check in gate["checks"]
            )
        )


if __name__ == "__main__":
    unittest.main()
