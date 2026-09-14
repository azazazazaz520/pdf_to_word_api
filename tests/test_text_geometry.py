"""字符级 PDF 版面模型的回归测试。"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from reportlab.pdfgen.canvas import Canvas

from src.export.content import _join_wrapped_lines
from src.ir.builder import build_document_ir
from src.layout.layout import extract_pdf_layout
from src.layout.models import _TextCharacter
from src.layout.text import (
    _build_text_line,
    _build_text_spans,
    _extract_text_characters,
)


def _write_geometry_sample(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(500.0, 500.0))
    canvas.setFont("Helvetica", 12)
    canvas.drawString(50, 450, "normal line")

    canvas.saveState()
    canvas.translate(100, 100)
    canvas.rotate(90)
    canvas.setFont("Courier-Bold", 11)
    canvas.drawString(0, 0, "rotated line")
    canvas.restoreState()

    canvas.saveState()
    canvas.translate(300, 300)
    canvas.rotate(45)
    canvas.setFont("Times-Italic", 10)
    canvas.drawString(0, 0, "diagonal")
    canvas.restoreState()
    canvas.save()


class TextGeometryTest(unittest.TestCase):
    @staticmethod
    def _character(
        text: str,
        *,
        x0: float,
        font_name: str,
        char_index: int,
    ) -> _TextCharacter:
        return _TextCharacter(
            text=text,
            x0=x0,
            top=10.0,
            x1=x0 + (0.5 if text.isspace() else 5.0),
            bottom=20.0,
            font_size=10.0,
            font_name=font_name,
            pdf_font_name=font_name,
            char_index=char_index,
        )

    def test_internal_character_indices_are_not_text_buffer_indices(self) -> None:
        class FakeTextPage:
            raw = object()

            def count_chars(self) -> int:
                return 4

            def get_charbox(self, index: int) -> tuple[float, float, float, float]:
                return {
                    0: (10.0, 10.0, 16.0, 20.0),
                    1: (16.0, 20.0, 16.0, 20.0),
                    2: (16.0, 10.0, 22.0, 20.0),
                    3: (22.0, 10.0, 28.0, 20.0),
                }[index]

        codepoints = (ord("A"), ord("\r"), ord("B"), ord("C"))
        with patch(
            "src.layout.text.pdfium_raw.FPDFText_GetUnicode",
            side_effect=lambda _raw, index: codepoints[index],
        ):
            characters = _extract_text_characters(FakeTextPage(), 100.0)

        self.assertEqual("".join(character.text for character in characters), "ABC")
        self.assertEqual(
            [character.char_index for character in characters],
            [0, 2, 3],
        )

    def test_glyphs_keep_content_geometry_font_and_direction(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "geometry.pdf"
            _write_geometry_sample(pdf_path)

            page = extract_pdf_layout(
                pdf_path,
                include_fidelity=True,
            ).pages[0]
            lines = {line.text: line for line in page.lines}
            text_blocks = list(page.text_blocks)

            self.assertEqual(set(lines), {"normal line", "rotated line", "diagonal"})
            self.assertEqual(
                sum(len(block.lines) for block in text_blocks),
                len(page.lines),
            )
            self.assertGreaterEqual(len(text_blocks), 3)
            for text, line in lines.items():
                self.assertEqual("".join(glyph.text for glyph in line.glyphs), text)
                self.assertEqual(len(line.glyphs), len(text))
                self.assertTrue(all(glyph.font_name for glyph in line.glyphs))
                self.assertTrue(all(glyph.font_size > 0 for glyph in line.glyphs))
                self.assertTrue(
                    all(glyph.char_index >= 0 for glyph in line.glyphs)
                )
                self.assertEqual(
                    [glyph.char_index for glyph in line.glyphs],
                    sorted(glyph.char_index for glyph in line.glyphs),
                )

            self.assertAlmostEqual(lines["normal line"].rotation, 0.0, places=1)
            self.assertAlmostEqual(lines["rotated line"].rotation, -90.0, places=1)
            self.assertAlmostEqual(lines["diagonal"].rotation, -45.0, places=1)
            self.assertAlmostEqual(lines["rotated line"].direction[1], -1.0, places=3)
            self.assertAlmostEqual(lines["diagonal"].direction[0], 0.707107, places=3)
            self.assertAlmostEqual(lines["diagonal"].direction[1], -0.707107, places=3)

    def test_default_route_keeps_raw_font_metadata(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "raw-fonts.pdf"
            _write_geometry_sample(pdf_path)

            page = extract_pdf_layout(pdf_path).pages[0]

            glyphs = [glyph for line in page.lines for glyph in line.glyphs]
            self.assertTrue(glyphs)
            self.assertTrue(all(glyph.font_name for glyph in glyphs))
            self.assertTrue(all(glyph.pdf_font_name for glyph in glyphs))
            self.assertTrue(all(glyph.font_size > 0 for glyph in glyphs))
            self.assertTrue(any(abs(glyph.rotation) > 1.0 for glyph in glyphs))

    def test_boundary_spaces_survive_font_span_aggregation(self) -> None:
        characters = [
            self._character("a", x0=0.0, font_name="Regular", char_index=0),
            self._character(" ", x0=5.5, font_name="Italic", char_index=1),
            self._character("b", x0=6.0, font_name="Italic", char_index=2),
        ]

        spans = _build_text_spans(characters, rotation=0.0)
        line = _build_text_line(characters, rotation=0.0)

        self.assertEqual([span.text for span in spans], ["a ", "b"])
        self.assertIsNotNone(line)
        self.assertEqual(line.text, "a b")
        self.assertEqual("".join(span.text for span in line.spans), "a b")
        self.assertEqual("".join(glyph.text for glyph in line.glyphs), "a b")

    def test_wrapped_line_hyphen_is_not_followed_by_extra_space(self) -> None:
        self.assertEqual(
            _join_wrapped_lines(["pre-", "trained"]),
            "pre-trained",
        )
        self.assertEqual(
            _join_wrapped_lines(["state-of-", "the-art"]),
            "state-of-the-art",
        )

    def test_same_visual_row_uses_left_to_right_order(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "reading-order.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(400.0, 400.0))
            canvas.setFont("Helvetica", 10)
            canvas.drawString(200, 300, "right")
            canvas.drawString(40, 300, "left")
            canvas.save()

            page = extract_pdf_layout(pdf_path).pages[0]
            row = [
                line.text
                for line in page.lines
                if line.text in {"left", "right"}
            ]

            self.assertEqual(row, ["left", "right"])

    def test_rotated_text_is_not_merged_into_horizontal_line(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "directions.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(400.0, 400.0))
            canvas.setFont("Helvetica", 10)
            canvas.drawString(40, 300, "horizontal")
            canvas.saveState()
            canvas.translate(180, 80)
            canvas.rotate(90)
            canvas.drawString(0, 0, "vertical")
            canvas.restoreState()
            canvas.save()

            page = extract_pdf_layout(
                pdf_path,
                include_fidelity=True,
            ).pages[0]
            texts = {line.text for line in page.lines}

            self.assertIn("horizontal", texts)
            self.assertIn("vertical", texts)
            self.assertEqual(
                {
                    round(line.rotation)
                    for line in page.lines
                    if line.text in {"horizontal", "vertical"}
                },
                {0, -90},
            )

    def test_glyphs_survive_ir_block_aggregation(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "ir-geometry.pdf"
            _write_geometry_sample(pdf_path)
            layout = extract_pdf_layout(pdf_path, include_fidelity=True)

            document = build_document_ir(
                source_pdf=pdf_path,
                page_routes=["text"],
                page_sizes=[(layout.pages[0].width, layout.pages[0].height)],
                layout=layout,
                fidelity=True,
            )
            ir_lines = [
                line
                for block in document.pages[0].blocks
                for line in block.lines
                if line.text in {"normal line", "rotated line", "diagonal"}
            ]

            self.assertEqual(len(ir_lines), 3)
            for line in ir_lines:
                self.assertEqual(
                    "".join(glyph.text for glyph in line.glyphs),
                    line.text,
                )
                self.assertEqual(
                    [glyph.char_index for glyph in line.glyphs],
                    sorted(glyph.char_index for glyph in line.glyphs),
                )

    def test_vertical_gap_separates_geometry_blocks(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "paragraphs.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(400.0, 400.0))
            canvas.setFont("Helvetica", 10)
            canvas.drawString(40, 320, "first paragraph line one")
            canvas.drawString(40, 308, "first paragraph line two")
            canvas.drawString(40, 270, "second paragraph")
            canvas.save()

            page = extract_pdf_layout(
                pdf_path,
                include_fidelity=True,
            ).pages[0]
            blocks = [
                block
                for block in page.text_blocks
                if block.text.strip()
            ]

            self.assertEqual(len(blocks), 2)
            self.assertEqual(len(blocks[0].lines), 2)
            self.assertEqual(blocks[1].text, "second paragraph")


if __name__ == "__main__":
    unittest.main()
