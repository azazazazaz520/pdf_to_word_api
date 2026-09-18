"""修复方案中新增的区域、坐标、归属和独立质量检查。"""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from reportlab.pdfgen.canvas import Canvas

from src.layout.models import PdfLayoutRegion, PdfTextLine
from src.layout.regions import (
    PageCoordinateTransform,
    merge_local_ocr_lines,
    normalize_model_regions,
)
from src.layout.text import _matrix_scale, _matrix_scales
from src.layout.vectors import _extract_vector_lines
from src.worker.quality import evaluate_final_content_quality
import pypdfium2 as pdfium


class RepairFeatureTest(unittest.TestCase):
    def test_font_matrix_uses_height_scale_and_keeps_non_uniform_components(self) -> None:
        self.assertAlmostEqual(_matrix_scale((3, 0, 0, 4, 20, 20)), 4.0)
        self.assertEqual(_matrix_scales((3, 0, 0, 4, 20, 20)), (3.0, 4.0))

    def test_model_bbox_is_converted_with_rotation(self) -> None:
        transform = PageCoordinateTransform(
            pixel_width=200,
            pixel_height=100,
            page_width=400,
            page_height=200,
        )
        self.assertEqual(transform.bbox_to_page((10, 20, 110, 70)), (20.0, 40.0, 220.0, 140.0))
        rotated = PageCoordinateTransform(
            pixel_width=200,
            pixel_height=100,
            page_width=400,
            page_height=200,
            rotation=90,
        )
        self.assertEqual(rotated.bbox_to_page((0, 0, 200, 100)), (0.0, 0.0, 400.0, 200.0))

    def test_transformed_path_lines_use_page_coordinates(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "transformed-lines.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(600, 600))
            canvas.saveState()
            canvas.translate(135, 100)
            canvas.rect(0, 0, 378, 40, stroke=1, fill=0)
            canvas.restoreState()
            canvas.save()

            document = pdfium.PdfDocument(str(pdf_path))
            try:
                horizontal, vertical = _extract_vector_lines(document[0], 600.0)
            finally:
                document.close()
            self.assertTrue(any(abs(line.x0 - 135.0) < 0.1 and abs(line.x1 - 513.0) < 0.1 for line in horizontal))
            self.assertTrue(any(abs(line.x - 135.0) < 0.1 for line in vertical))

    def test_model_regions_keep_page_source_and_order_fields(self) -> None:
        line = PdfTextLine(
            text="body",
            x0=20,
            top=20,
            x1=80,
            bottom=32,
            font_size=10,
            source_char_ids=("1:4",),
        )
        regions = normalize_model_regions(
            [{"bbox": [10, 10, 100, 50], "label": "text", "score": 0.8, "order_hint": 2}],
            page_number=1,
            transform=PageCoordinateTransform(100, 100, 200, 200),
            lines=[line],
        )
        self.assertEqual(regions[0].source, "layout-model")
        self.assertEqual(regions[0].order_hint, 2)
        self.assertEqual(regions[0].source_char_ids, ("1:4",))

    def test_local_ocr_does_not_duplicate_usable_native_text(self) -> None:
        region = PdfLayoutRegion("r1", 1, "body", (0, 0, 100, 100))
        native = PdfTextLine("Hello", 10, 10, 50, 20, 10)
        ocr = PdfTextLine("Hello", 10, 10, 50, 20, 10)
        merged, report = merge_local_ocr_lines([native], [ocr], region=region)
        self.assertEqual([line.text for line in merged], ["Hello"])
        self.assertEqual(report["source"], "native")

    def test_final_quality_gate_reports_missing_and_duplicate_source_chars(self) -> None:
        report = evaluate_final_content_quality(
            source_char_ids=("1:1", "1:2"),
            placements=[
                {"source_char_ids": ["1:1", "1:1"]},
            ],
            source_text="AB",
            output_text="A",
        )
        self.assertEqual(report["status"], "failed")
        self.assertEqual(
            {check["name"] for check in report["checks"] if check["status"] == "failed"},
            {
                "source_character_coverage",
                "source_character_unique_placement",
                "text_content_presence",
            },
        )


if __name__ == "__main__":
    unittest.main()
