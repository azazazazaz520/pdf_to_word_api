from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from src.layout.layout import extract_pdf_layout


class PdfLayoutFeaturesTest(unittest.TestCase):
    def test_two_column_lines_are_read_column_by_column(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "two-column.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(600, 800))
            for index, y in enumerate((700, 680, 660, 640, 620, 600)):
                canvas.drawString(72, y, f"Left {index} short line")
                canvas.drawString(360, y, f"Right {index} short line")
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]
            texts = [line.text for line in page.lines]

            self.assertTrue(page.columns)
            self.assertEqual(page.tables, ())
            self.assertLess(
                texts.index("Left 5 short line"),
                texts.index("Right 0 short line"),
            )
            self.assertLess(
                texts.index("Left 0 short line"),
                texts.index("Left 5 short line"),
            )

    def test_repeated_header_footer_is_marked_and_filtered(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "headers.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(360, 480))
            for page_number in (1, 2):
                canvas.setFont("Helvetica", 9)
                canvas.drawString(36, 455, "Quarterly Report - Internal")
                canvas.drawString(36, 30, f"Page {page_number}")
                canvas.setFont("Helvetica", 11)
                canvas.drawString(
                    36,
                    400,
                    f"Body text page {page_number} with content.",
                )
                canvas.showPage()
            canvas.save()

            layout = extract_pdf_layout(pdf_path)

            header_footer_flags = [
                (line.text, line.is_header_footer)
                for page in layout.pages
                for line in page.lines
            ]
            self.assertIn(("Quarterly Report - Internal", True), header_footer_flags)
            self.assertIn(("Page 1", True), header_footer_flags)
            self.assertIn(("Page 2", True), header_footer_flags)
            self.assertIn(("Body text page 1 with content.", False), header_footer_flags)

    def test_embedded_image_is_extracted_with_bbox(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "embedded.png"
            pdf_path = root / "embedded-image.pdf"
            Image.new("RGB", (200, 100), color=(20, 80, 160)).save(image_path)

            canvas = Canvas(str(pdf_path), pagesize=(400, 300))
            canvas.drawImage(
                ImageReader(str(image_path)),
                50,
                50,
                width=200,
                height=100,
            )
            canvas.drawString(50, 270, "Text on the same page")
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]

            self.assertEqual(len(page.images), 1)
            image_block = page.images[0]
            self.assertTrue(image_block.data.startswith(b"\x89PNG"))
            self.assertAlmostEqual(image_block.bbox[0], 50.0, places=1)
            self.assertAlmostEqual(image_block.bbox[1], 150.0, places=1)
            self.assertAlmostEqual(image_block.bbox[2], 250.0, places=1)
            self.assertAlmostEqual(image_block.bbox[3], 250.0, places=1)


    def test_embedded_image_respects_pixel_limit(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "large.png"
            pdf_path = root / "large-image.pdf"
            Image.new("RGB", (400, 300), color=(40, 90, 150)).save(image_path)

            canvas = Canvas(str(pdf_path), pagesize=(400, 300))
            canvas.drawImage(
                ImageReader(str(image_path)),
                20,
                20,
                width=360,
                height=260,
            )
            canvas.save()

            page = extract_pdf_layout(
                pdf_path,
                image_max_pixels=10_000,
            ).pages[0]

            self.assertEqual(len(page.images), 1)
            image_block = page.images[0]
            with Image.open(BytesIO(image_block.data)) as encoded:
                self.assertLessEqual(encoded.width * encoded.height, 10_000)
            self.assertIn(image_block.mime_type, {"image/png", "image/jpeg"})

    def test_original_jpeg_stream_is_passed_through(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "photo.jpg"
            pdf_path = root / "photo.pdf"
            Image.new("RGB", (640, 480), color=(120, 80, 40)).save(
                image_path,
                format="JPEG",
                quality=85,
            )

            canvas = Canvas(str(pdf_path), pagesize=(320, 240))
            canvas.drawImage(
                ImageReader(str(image_path)),
                20,
                20,
                width=280,
                height=200,
            )
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]

            self.assertEqual(len(page.images), 1)
            image_block = page.images[0]
            self.assertEqual(image_block.mime_type, "image/jpeg")
            self.assertTrue(image_block.data.startswith(b"\xff\xd8"))
            self.assertTrue(image_block.data.endswith(b"\xff\xd9"))

    def test_image_extraction_can_be_limited_to_selected_pages(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "page-image.png"
            pdf_path = root / "two-pages.pdf"
            Image.new("RGB", (120, 80), color=(200, 200, 200)).save(image_path)

            canvas = Canvas(str(pdf_path), pagesize=(320, 240))
            canvas.drawImage(
                ImageReader(str(image_path)),
                20,
                20,
                width=140,
                height=100,
            )
            canvas.showPage()
            canvas.drawImage(
                ImageReader(str(image_path)),
                20,
                20,
                width=140,
                height=100,
            )
            canvas.save()

            layout = extract_pdf_layout(pdf_path, include_page_images={1})

            self.assertEqual(layout.pages[0].images, ())
            self.assertEqual(len(layout.pages[1].images), 1)



    def test_borderless_table_is_detected(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "borderless-table.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(420, 300))
            rows = (
                ("Name", "Value", "Unit"),
                ("Alpha", "12", "V"),
                ("Beta", "24", "mA"),
                ("Gamma", "36", "ms"),
            )
            for row_index, values in enumerate(rows):
                y = 240 - row_index * 40
                for x, value in zip((40, 170, 300), values):
                    canvas.drawString(x, y, value)
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]

            self.assertEqual(len(page.tables), 1)
            table = page.tables[0]
            self.assertEqual(table.column_count, 3)
            self.assertGreaterEqual(table.row_count, 4)
            self.assertEqual(table.rows[0], ("Name", "Value", "Unit"))
            self.assertEqual(table.rows[1], ("Alpha", "12", "V"))



    def test_two_column_borderless_table_with_caption_is_detected(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "two-column-table.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(420, 300))
            canvas.drawString(40, 265, "表 1 参数")
            rows = (
                ("Name", "Value"),
                ("Alpha", "12"),
                ("Beta", "24"),
                ("Gamma", "36"),
            )
            for row_index, values in enumerate(rows):
                y = 235 - row_index * 35
                canvas.drawString(40, y, values[0])
                canvas.drawString(220, y, values[1])
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]

            self.assertEqual(len(page.tables), 1)
            table = page.tables[0]
            self.assertEqual(table.column_count, 2)
            self.assertEqual(table.rows[0], ("Name", "Value"))
            self.assertEqual(table.rows[1], ("Alpha", "12"))

    def test_borderless_table_infers_column_span(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "merged-borderless.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(420, 320))
            canvas.drawString(
                40,
                270,
                "Merged header across all columns and more text",
            )
            rows = (
                ("Name", "Value", "Unit"),
                ("Alpha", "12", "V"),
                ("Beta", "24", "mA"),
            )
            for row_index, values in enumerate(rows):
                y = 230 - row_index * 35
                for x, value in zip((40, 170, 300), values):
                    canvas.drawString(x, y, value)
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]

            self.assertEqual(len(page.tables), 1)
            table = page.tables[0]
            self.assertTrue(
                any(cell.column_span == 3 for cell in table.cells)
            )


if __name__ == "__main__":
    unittest.main()
