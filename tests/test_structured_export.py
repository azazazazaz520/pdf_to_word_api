from __future__ import annotations

import unittest
import zipfile
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from PIL import Image

from src.export.structured import export_structured_docx
from src.ir.model import IRBlock, IRDocument, IRPage
from src.layout.models import PdfTable, PdfTableCell


def _image_bytes() -> bytes:
    output = BytesIO()
    Image.new("RGB", (120, 60), color=(40, 100, 180)).save(output, format="PNG")
    return output.getvalue()


def _pdf_table() -> PdfTable:
    return PdfTable(
        bbox=(40.0, 70.0, 320.0, 130.0),
        column_boundaries=(40.0, 180.0, 320.0),
        row_boundaries=(70.0, 100.0, 130.0),
        rows=(("表头 A", "表头 B"), ("单元格 1", "单元格 2")),
        cells=(
            PdfTableCell(
                row_index=0,
                column_index=0,
                row_span=1,
                column_span=1,
                bbox=(40.0, 70.0, 180.0, 100.0),
                text="表头 A",
                font_size=9.0,
                bold=True,
            ),
            PdfTableCell(
                row_index=0,
                column_index=1,
                row_span=1,
                column_span=1,
                bbox=(180.0, 70.0, 320.0, 100.0),
                text="表头 B",
                font_size=9.0,
                bold=True,
            ),
            PdfTableCell(
                row_index=1,
                column_index=0,
                row_span=1,
                column_span=1,
                bbox=(40.0, 100.0, 180.0, 130.0),
                text="单元格 1",
                font_size=9.0,
            ),
            PdfTableCell(
                row_index=1,
                column_index=1,
                row_span=1,
                column_span=1,
                bbox=(180.0, 100.0, 320.0, 130.0),
                text="单元格 2",
                font_size=9.0,
            ),
        ),
    )


class StructuredExportTest(unittest.TestCase):
    def test_writes_flow_text_table_omml_and_inline_image(self) -> None:
        ir = IRDocument(
            title="结构化示例",
            metadata={"outline": [{"page": 1, "title": "Chapter 1"}]},
            pages=[
                IRPage(
                    page_number=1,
                    width=400.0,
                    height=300.0,
                    route="text",
                    blocks=[
                        IRBlock(
                            kind="paragraph",
                            page=1,
                            text="普通正文应当保持为可重排的 Word 段落。",
                        ),
                        IRBlock(
                            kind="table",
                            page=1,
                            bbox=(40.0, 70.0, 320.0, 130.0),
                            table=_pdf_table(),
                        ),
                        IRBlock(
                            kind="formula",
                            page=1,
                            text=r"x = \frac{a}{b} + 1",
                            bbox=(40.0, 140.0, 320.0, 170.0),
                        ),
                        IRBlock(
                            kind="image",
                            page=1,
                            bbox=(40.0, 180.0, 160.0, 240.0),
                            image_bytes=_image_bytes(),
                            image_width=120.0,
                            image_height=60.0,
                            image_alt="示例图片",
                        ),
                    ],
                )
            ],
        )

        with TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "result.docx"
            report = export_structured_docx(
                ir,
                output_path,
                include_toc=True,
                include_bookmarks=True,
            )

            self.assertEqual(report["text_paragraph_count"], 1)
            self.assertEqual(report["table_count"], 1)
            self.assertEqual(report["formula_count"], 1)
            self.assertEqual(report["formula_omml_count"], 1)
            self.assertEqual(report["image_count"], 1)
            document = Document(output_path)
            self.assertEqual(len(document.tables), 1)
            self.assertIn(
                "普通正文应当保持为可重排的 Word 段落。",
                "\n".join(paragraph.text for paragraph in document.paragraphs),
            )
            self.assertEqual(document.tables[0].cell(1, 1).text, "单元格 2")
            with zipfile.ZipFile(output_path) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertNotIn("<w:framePr", document_xml)
            self.assertIn("<w:tbl", document_xml)
            self.assertIn("<m:oMath", document_xml)
            self.assertIn("<wp:inline", document_xml)
            self.assertIn("w:bookmarkStart", document_xml)
            self.assertIn("Chapter_1", document_xml)
            self.assertIn("TOC", document_xml)

    def test_writes_ocr_html_table_as_word_table(self) -> None:
        ir = IRDocument(
            pages=[
                IRPage(
                    page_number=1,
                    width=400.0,
                    height=300.0,
                    route="ocr",
                    blocks=[
                        IRBlock(
                            kind="html_table",
                            page=1,
                            text=(
                                "<table><tr><th colspan='2'>总览</th></tr>"
                                "<tr><td rowspan='2'>A</td><td>B</td></tr>"
                                "<tr><td>C</td></tr></table>"
                            ),
                        )
                    ],
                )
            ]
        )

        with TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "ocr-table.docx"
            report = export_structured_docx(ir, output_path)

            self.assertEqual(report["table_count"], 1)
            self.assertEqual(report["html_table_count"], 1)
            document = Document(output_path)
            self.assertEqual(len(document.tables), 1)
            self.assertEqual(document.tables[0].cell(0, 0).text, "总览")
            self.assertEqual(document.tables[0].cell(1, 0).text, "A")
            self.assertEqual(document.tables[0].cell(2, 1).text, "C")

    def test_bad_table_is_degraded_locally(self) -> None:
        ir = IRDocument(
            pages=[
                IRPage(
                    page_number=1,
                    width=400.0,
                    height=300.0,
                    route="text",
                    blocks=[IRBlock(kind="table", page=1, table=object())],
                )
            ]
        )

        with TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "bad-table.docx"
            report = export_structured_docx(ir, output_path)

            self.assertTrue(output_path.is_file())
            self.assertEqual(report["table_native_error_count"], 1)
            self.assertEqual(report["skipped_block_count"], 1)
            self.assertEqual(report["placements"][0]["status"], "skipped")

    def test_dense_vectors_are_compacted_without_per_object_rendering(self) -> None:
        ir = IRDocument(
            pages=[
                IRPage(
                    page_number=1,
                    width=400.0,
                    height=300.0,
                    route="text",
                    blocks=[
                        IRBlock(
                            kind="vector",
                            page=1,
                            bbox=(20.0 + index, 40.0, 21.0 + index, 41.0),
                        )
                        for index in range(65)
                    ],
                )
            ]
        )

        with TemporaryDirectory() as temporary_directory:
            output_path = Path(temporary_directory) / "dense-vectors.docx"
            report = export_structured_docx(ir, output_path)

            self.assertEqual(report["vector_compacted_page_count"], 1)
            self.assertEqual(report["vector_skipped_count"], 65)
            self.assertEqual(report["vector_image_fallback_count"], 0)
            self.assertTrue(output_path.is_file())


if __name__ == "__main__":
    unittest.main()
