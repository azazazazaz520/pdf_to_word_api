from __future__ import annotations

from io import BytesIO
import zipfile
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from docx.oxml.ns import qn
from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas

from src.layout.layout import extract_pdf_layout
from src.export.content import _group_text_lines
from src.export.document_setup import _set_document_styles


class HybridExportTest(unittest.TestCase):
    def test_text_grouping_joins_wrapped_lines_without_extra_spaces(self) -> None:
        blocks = _group_text_lines(
            "标题\n第一行内容\n第二行内容。\n1. 第一项\n续行。",
            is_document_start=True,
        )

        self.assertEqual(
            blocks,
            [
                ("title", "标题"),
                ("body", "第一行内容第二行内容。"),
                ("ordered", "第一项续行。"),
            ],
        )

    def test_list_item_colon_keeps_following_explanation_in_same_item(self) -> None:
        blocks = _group_text_lines(
            "四、总结\n1. 第一项：\n这是第一项的说明。\n2. 第二项：\n这是第二项的说明。"
        )

        self.assertEqual(
            blocks,
            [
                ("heading1", "四、总结"),
                ("ordered", "第一项：这是第一项的说明。"),
                ("ordered", "第二项：这是第二项的说明。"),
            ],
        )

    def test_common_parenthesized_numbers_are_ordered_items(self) -> None:
        blocks = _group_text_lines("（1）第一项。\n(2) 第二项。")

        self.assertEqual(
            blocks,
            [("ordered", "第一项。"), ("ordered", "第二项。")],
        )

    def test_common_math_operator_line_is_a_formula(self) -> None:
        blocks = _group_text_lines("目标函数：\n∑ᵢ xᵢ = 1")

        self.assertEqual(
            blocks,
            [("body", "目标函数："), ("formula", "∑ᵢ xᵢ = 1")],
        )

if __name__ == "__main__":
    unittest.main()
