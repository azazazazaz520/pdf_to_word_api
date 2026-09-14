from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas

from src.pdf_routing import (
    _count_garbled_characters,
    _read_page_text,
    analyze_pdf_text,
    normalize_page_text,
)

_WINGDINGS_FONT_PATH = Path(r"C:\Windows\Fonts\wingding.ttf")
_WINGDINGS_FONT_NAME = "WingdingsRoutingTest"
# 符号字体私有区中未被归一化覆盖的码位，抽取后仍保持私有区类别
_UNMAPPED_SYMBOL_CHARACTER = "\uf0a8"


def _create_text_pdf(path: Path) -> None:
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Fast route page one with selectable text")
    canvas.showPage()
    canvas.drawString(36, 200, "Fast route page two with selectable text")
    canvas.save()


def _create_image_backed_text_pdf(root: Path, path: Path) -> None:
    background_path = root / "background.png"
    Image.new("RGB", (720, 480), color=(220, 230, 240)).save(background_path)
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawImage(ImageReader(str(background_path)), 0, 0, width=360, height=240)
    canvas.drawString(36, 200, "Only a partial selectable text layer")
    canvas.save()


def _create_symbol_font_pdf(path: Path, *, garbled_count: int) -> None:
    """写出可读字符数恒定、异常字符数由参数决定的页面。

    可读文本两端不加空白，避免抽取时首尾空格被折叠导致字符总数随参数变化；
    恒定写入 2 个未映射符号字符作为底噪，使字符总数只随异常字符数增长。
    """
    if _WINGDINGS_FONT_NAME not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(_WINGDINGS_FONT_NAME, str(_WINGDINGS_FONT_PATH)))
    canvas = Canvas(str(path), pagesize=(360, 240))
    canvas.drawString(36, 200, "Selectable|text|content|on|this|page" * 4)
    canvas.setFont(_WINGDINGS_FONT_NAME, 12)
    canvas.drawString(
        36,
        160,
        _UNMAPPED_SYMBOL_CHARACTER * max(garbled_count, 2),
    )
    canvas.save()


class PdfRoutingTest(unittest.TestCase):
    def test_compatibility_characters_are_normalized_for_detection(self) -> None:
        # 判定路径折叠全部兼容字符；版面文本使用的共享函数只折叠数学字母
        value = "𝑥 + 𝑦 + 𝛼 + ⼀⽂⽤⽅ + ﬁ 项目"

        self.assertEqual(_read_page_text(value), "x + y + α + 一文用方 + fi 项目")
        self.assertEqual(normalize_page_text(value), "x + y + α + ⼀⽂⽤⽅ + ﬁ 项目")

    def test_symbol_font_private_use_characters_are_normalized(self) -> None:
        value = "\uf06c 主控芯片\uf071 默认连接\uf0b7 提示"

        self.assertEqual(normalize_page_text(value), "• 主控芯片❑ 默认连接• 提示")
        self.assertEqual(
            _count_garbled_characters(normalize_page_text(value)),
            0,
        )

    def test_rare_scripts_and_radicals_are_not_treated_as_garbled(self) -> None:
        # 康熙部首为正常中文兼容字符，奥里亚文与埃塞俄比亚文为正常文字系统
        value = "⼀⽂⽤⽅ ଵଶ ቐ ሀ"

        self.assertEqual(_count_garbled_characters(value), 0)

    def test_unmapped_private_use_characters_are_treated_as_garbled(self) -> None:
        self.assertEqual(_count_garbled_characters("正常文本\ue000\ufffd"), 2)

    def test_garbled_penalty_scales_with_ratio(self) -> None:
        if not _WINGDINGS_FONT_PATH.exists():
            self.skipTest("缺少可用于构造私有区字符的符号字体")
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)

            def analyze(name: str, garbled_count: int):
                path = root / f"{name}.pdf"
                _create_symbol_font_pdf(path, garbled_count=garbled_count)
                return analyze_pdf_text(path, min_page_chars=20).pages[0]

            few = analyze("few", 2)
            medium = analyze("medium", 5)
            many = analyze("many", 20)

            base = few.text_char_count
            self.assertEqual((few.garbled_char_count, medium.garbled_char_count), (2, 5))
            self.assertEqual(many.garbled_char_count, 20)
            self.assertEqual(medium.text_char_count - base, 3)
            self.assertEqual(many.text_char_count - base, 18)

            # 同底噪下，异常字符越多扣分越多
            self.assertGreater(few.quality_score, medium.quality_score)
            self.assertGreater(medium.quality_score, many.quality_score)

            # 占比未超过阈值时只按比例扣分，页面仍判为高质量
            self.assertAlmostEqual(
                medium.quality_score,
                1.0 - 0.35 * (5 / medium.text_char_count) / 0.05,
                places=3,
            )
            self.assertTrue(medium.is_high_quality(min_page_chars=20))

            # 占比超过 5% 阈值时扣满 0.35，页面不再判为高质量
            self.assertGreater(many.garbled_char_count / many.text_char_count, 0.05)
            self.assertAlmostEqual(many.quality_score, 0.65, places=3)
            self.assertFalse(many.is_high_quality(min_page_chars=20))

    def test_analyze_pdf_text_detects_text_layer(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "text.pdf"
            _create_text_pdf(pdf_path)

            analysis = analyze_pdf_text(pdf_path, min_page_chars=20)

            self.assertEqual(analysis.page_count, 2)
            self.assertEqual(analysis.usable_page_count, 2)
            self.assertEqual(analysis.usable_page_ratio, 1.0)
            self.assertIn("Fast route page one", analysis.page_texts[0])
            self.assertTrue(
                analysis.has_usable_text_layer(
                    min_page_chars=20,
                    min_page_ratio=0.6,
                )
            )
            self.assertEqual(analysis.high_quality_page_count, 2)
            self.assertTrue(
                analysis.has_complete_text_layer(
                    min_page_chars=20,
                    min_high_quality_ratio=0.8,
                )
            )

    def test_image_backed_text_layer_is_not_complete(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "image-backed.pdf"
            _create_image_backed_text_pdf(root, pdf_path)

            analysis = analyze_pdf_text(pdf_path, min_page_chars=20)

            self.assertEqual(analysis.page_count, 1)
            self.assertEqual(analysis.usable_page_count, 1)
            self.assertEqual(analysis.full_page_image_page_count, 1)
            self.assertEqual(analysis.high_quality_page_count, 0)
            self.assertTrue(
                analysis.has_usable_text_layer(
                    min_page_chars=20,
                    min_page_ratio=0.6,
                )
            )
            self.assertFalse(
                analysis.has_complete_text_layer(
                    min_page_chars=20,
                    min_high_quality_ratio=0.8,
                )
            )

if __name__ == "__main__":
    unittest.main()
