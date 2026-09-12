from __future__ import annotations

import unittest

from src.ir.model import IRBlock, IRDocument, IRPage, IRTextLine, IRWarning


class DocumentIRTest(unittest.TestCase):
    def test_quality_report_counts_blocks_and_review_pages(self) -> None:
        document = IRDocument(
            pages=[
                IRPage(
                    page_number=1,
                    width=595.0,
                    height=842.0,
                    route="text",
                    blocks=[
                        IRBlock(kind="heading1", page=1, text="标题"),
                        IRBlock(kind="paragraph", page=1, text="正文"),
                        IRBlock(kind="table", page=1, table=object()),
                        IRBlock(
                            kind="image",
                            page=1,
                            image_bytes=b"image",
                            image_width=100.0,
                            image_height=80.0,
                        ),
                    ],
                ),
                IRPage(
                    page_number=2,
                    width=595.0,
                    height=842.0,
                    route="page_image",
                    editable=False,
                    blocks=[
                        IRBlock(
                            kind="page_image",
                            page=2,
                            image_bytes=b"page",
                            image_width=595.0,
                            image_height=842.0,
                        )
                    ],
                    warnings=[
                        IRWarning(
                            code="page_preserved_as_image",
                            message="页面以图片保留",
                            page=2,
                            severity="warning",
                        )
                    ],
                ),
            ]
        )

        report = document.quality_report()

        self.assertEqual(report["route_summary"], {"text": 1, "page_image": 1})
        self.assertEqual(report["table_count"], 1)
        self.assertEqual(report["media_count"], 2)
        self.assertEqual(report["heading_count"], 1)
        self.assertEqual(report["editable_page_count"], 1)
        self.assertEqual(report["visual_only_page_count"], 1)
        self.assertEqual(report["needs_review_pages"], [2])
        self.assertEqual(report["page_results"][0]["route"], "text")
        self.assertEqual(report["page_results"][1]["media_count"], 1)

    def test_route_summary_defaults_to_empty(self) -> None:
        report = IRDocument().quality_report()

        self.assertEqual(report["route_summary"], {})
        self.assertEqual(report["source_page_count"], 0)
        self.assertEqual(report["needs_review_pages"], [])

    def test_compact_quality_report_aggregates_info_warnings(self) -> None:
        pages = []
        for page_number in range(1, 251):
            warnings = [
                IRWarning(
                    code="header_footer_filtered",
                    message="已过滤 1 行页眉页脚",
                    page=page_number,
                    severity="info",
                )
            ]
            if page_number == 10:
                warnings.append(
                    IRWarning(
                        code="page_preserved_as_image",
                        message="页面以图片保留",
                        page=page_number,
                        severity="warning",
                    )
                )
            pages.append(
                IRPage(
                    page_number=page_number,
                    width=595.0,
                    height=842.0,
                    route="text",
                    blocks=[
                        IRBlock(kind="paragraph", page=page_number, text="内容")
                    ],
                    warnings=warnings,
                )
            )

        report = IRDocument(pages=pages).quality_report(compact=True)

        self.assertTrue(report["page_results_compacted"])
        self.assertEqual(report["page_results_total"], 250)
        self.assertEqual([item["page"] for item in report["page_results"]], [10])
        header_warning = next(
            item
            for item in report["warnings"]
            if item["code"] == "header_footer_filtered"
        )
        self.assertEqual(header_warning["page_count"], 250)
        self.assertEqual(header_warning["pages"], "1-250")
        self.assertIn(10, report["needs_review_pages"])


    def test_fidelity_report_aggregates_fallback_regions(self) -> None:
        block = IRBlock(
            kind="paragraph",
            page=1,
            text="低置信度文本",
            bbox=(72.0, 100.0, 300.0, 112.0),
            confidence=0.35,
            lines=(
                IRTextLine(
                    text="低置信度文本",
                    bbox=(72.0, 100.0, 300.0, 112.0),
                    font_name="Arial",
                    font_size=11.0,
                    color=(0, 0, 0),
                    alignment="left",
                    z_order=3,
                ),
            ),
            fallback_reason="low_confidence:0.350",
        )
        header = IRBlock(
            kind="paragraph",
            page=1,
            text="页眉",
            layer="header",
            bbox=(72.0, 30.0, 200.0, 40.0),
        )
        page = IRPage(
            page_number=1,
            width=595.0,
            height=842.0,
            route="text",
            blocks=[block, header],
            reconstruction_confidence=0.9,
            fidelity={
                "rebuild_confidence": 0.9,
                "fallback_regions": [
                    {
                        "kind": "paragraph",
                        "layer": "body",
                        "bbox": [72.0, 100.0, 300.0, 112.0],
                        "reason": "low_confidence:0.350",
                    }
                ],
                "bbox_error_max": 0.4,
                "font_size_error_max": 0.05,
            },
        )
        document = IRDocument(pages=[page], metadata={"export_mode": "fidelity"})

        report = document.fidelity_report()

        self.assertEqual(report["page_count"], 1)
        self.assertEqual(report["fallback_region_count"], 1)
        self.assertEqual(report["fallback_pages"], [1])
        self.assertEqual(report["mean_rebuild_confidence"], 0.9)
        self.assertEqual(report["max_bbox_error"], 0.4)
        self.assertEqual(report["max_font_size_error"], 0.05)
        self.assertEqual(len(report["unreconstructable_regions"]), 1)
        self.assertEqual(report["unreconstructable_regions"][0]["page"], 1)

    def test_page_layer_helpers_split_header_footer_and_body(self) -> None:
        page = IRPage(
            page_number=1,
            width=595.0,
            height=842.0,
            route="text",
            blocks=[
                IRBlock(kind="paragraph", page=1, text="正文"),
                IRBlock(kind="paragraph", page=1, text="页眉", layer="header"),
                IRBlock(
                    kind="image",
                    page=1,
                    layer="footer",
                    image_bytes=b"logo",
                    fallback_reason="explicit_fallback_image",
                ),
            ],
        )

        self.assertEqual(len(page.body_blocks), 1)
        self.assertEqual(len(page.header_blocks), 1)
        self.assertEqual(len(page.footer_blocks), 1)
        self.assertEqual(len(page.fallback_blocks), 1)
        self.assertTrue(page.blocks[2].has_fallback)
        self.assertEqual(len(page.fallback_regions), 1)

    def test_quality_report_includes_fidelity_section(self) -> None:
        page = IRPage(
            page_number=1,
            width=595.0,
            height=842.0,
            route="text",
            blocks=[IRBlock(kind="paragraph", page=1, text="内容")],
            reconstruction_confidence=0.95,
            fidelity={"rebuild_confidence": 0.95, "fallback_regions": []},
        )
        document = IRDocument(
            pages=[page],
            metadata={"export_mode": "fidelity_hybrid"},
        )

        report = document.quality_report()

        self.assertIsNotNone(report["fidelity"])
        self.assertEqual(report["fidelity"]["mode"], "fidelity_hybrid")
        self.assertEqual(report["page_results"][0]["rebuild_confidence"], 0.95)
        self.assertEqual(report["text_line_count"], 0)


if __name__ == "__main__":
    unittest.main()
