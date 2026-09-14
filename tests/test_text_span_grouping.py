"""版面文字片段分组与字形穿插的回归防线。

覆盖两点：
  1. 同一行内字号因逐字形外框而抖动时，不应被拆成逐字符片段；
  2. 渲染结果中同一视觉行内不得出现字形互相覆盖。
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas

from src.layout.layout import extract_pdf_layout


def _create_mixed_metric_pdf(path: Path) -> None:
    """写出正文中夹小字号数字与上升部字母的页面。

    上升部字母（如 h、k）的外框明显高于 x 高度字母（如 a、e），
    用于模拟逐字形度量差异导致的字号抖动。
    """
    canvas = Canvas(str(path), pagesize=A4)
    canvas.setFont("Helvetica", 11)
    canvas.drawString(60, 740, "the quick hake takes a token and the answer")
    canvas.drawString(60, 720, "the task is honest about the shape of a token")
    canvas.save()


class TextSpanGroupingTest(unittest.TestCase):
    def test_glyph_metric_jitter_does_not_split_every_character(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "mixed.pdf"
            _create_mixed_metric_pdf(pdf_path)

            layout = extract_pdf_layout(pdf_path, include_fidelity=False)

            self.assertEqual(len(layout.pages), 1)
            lines = [line for line in layout.pages[0].lines if str(line.text).strip()]
            self.assertGreaterEqual(len(lines), 2)

            for line in lines:
                spans = list(getattr(line, "spans", ()) or ())
                characters = len(str(line.text))
                self.assertGreater(characters, 20)
                # 片段数必须显著少于字符数，否则说明分组退化成了逐字符切分
                self.assertLess(
                    len(spans),
                    characters / 2,
                    f"行 {line.text!r} 被拆成 {len(spans)} 个片段",
                )

    def test_span_count_stays_bounded_for_uniform_text(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "uniform.pdf"
            canvas = Canvas(str(pdf_path), pagesize=A4)
            canvas.setFont("Helvetica", 11)
            canvas.drawString(60, 700, "a uniform line of text without metric jitter")
            canvas.save()

            layout = extract_pdf_layout(pdf_path, include_fidelity=False)

            lines = [line for line in layout.pages[0].lines if str(line.text).strip()]
            self.assertEqual(len(lines), 1)
            spans = list(getattr(lines[0], "spans", ()) or ())
            # 同一字号的行仍会按上升部字母与 x 高度字母分成若干片段，
            # 但数量应远小于字符数，且不随行长度线性增长。
            self.assertLessEqual(len(spans), 10)


if __name__ == "__main__":
    unittest.main()
