from __future__ import annotations

import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from reportlab.pdfgen.canvas import Canvas

from src.ir.model import IRBlock, IRDocument, IRPage
from src.ir.outline import apply_outline_to_headings, extract_pdf_outline


class PdfOutlineTest(unittest.TestCase):
    def test_extract_pdf_outline_returns_titles_pages_and_levels(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "outline.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(360, 240))
            canvas.bookmarkPage("chapter1")
            canvas.addOutlineEntry("Chapter 1", "chapter1", level=0)
            canvas.drawString(36, 200, "Chapter one")
            canvas.showPage()
            canvas.bookmarkPage("chapter2")
            canvas.addOutlineEntry("Chapter 2", "chapter2", level=1)
            canvas.drawString(36, 200, "Chapter two")
            canvas.save()

            entries = extract_pdf_outline(pdf_path)

            self.assertEqual(
                entries,
                [
                    {"title": "Chapter 1", "page": 1, "level": 0},
                    {"title": "Chapter 2", "page": 2, "level": 1},
                ],
            )

    def test_outline_levels_are_applied_to_matching_headings(self) -> None:
        document = IRDocument(
            pages=[
                IRPage(
                    page_number=1,
                    width=595.0,
                    height=842.0,
                    route="text",
                    blocks=[
                        IRBlock(kind="heading1", page=1, text="第1章"),
                        IRBlock(kind="paragraph", page=1, text="如何使用本书"),
                        IRBlock(
                            kind="heading2",
                            page=1,
                            text="1.1 本书的学习顺序",
                        ),
                    ],
                )
            ]
        )
        outline = [
            {"title": "第1章 如何使用本书", "page": 1, "level": 0},
            {"title": "1.1 本书的学习顺序", "page": 1, "level": 1},
        ]

        matched = apply_outline_to_headings(document, outline)

        self.assertGreaterEqual(matched, 2)
        self.assertEqual(document.pages[0].blocks[0].kind, "heading1")
        self.assertEqual(document.pages[0].blocks[0].level, 1)
        self.assertEqual(document.pages[0].blocks[2].kind, "heading2")
        self.assertEqual(document.pages[0].blocks[2].level, 2)
        self.assertEqual(
            document.pages[0].blocks[0].meta.get("outline_title"),
            "第1章 如何使用本书",
        )



if __name__ == "__main__":
    unittest.main()