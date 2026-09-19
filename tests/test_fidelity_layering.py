from __future__ import annotations

import unittest

from src.export.fidelity import (
    _can_place_header_footer_item,
    _filled_vector_is_behind_text,
)
from src.ir.model import IRBlock, IRPage
from src.layout.models import PdfVectorObject


def _page_with_text() -> IRPage:
    return IRPage(
        page_number=1,
        width=100.0,
        height=100.0,
        route="text",
        blocks=[
            IRBlock(
                kind="paragraph",
                page=1,
                text="字段标签",
                bbox=(20.0, 20.0, 80.0, 30.0),
                z_order=5,
            )
        ],
    )


class FidelityLayeringTest(unittest.TestCase):
    def test_header_footer_media_placement_checks_are_reachable(self) -> None:
        image = IRBlock(kind="image", page=1, image_bytes=b"image")
        vector = IRBlock(
            kind="vector",
            page=1,
            bbox=(0.0, 0.0, 10.0, 10.0),
            vector=PdfVectorObject(
                kind="line",
                bbox=(0.0, 0.0, 10.0, 10.0),
            ),
        )

        self.assertTrue(_can_place_header_footer_item(image))
        self.assertTrue(_can_place_header_footer_item(vector))
        self.assertFalse(_can_place_header_footer_item(IRBlock(kind="image", page=1)))

    def test_fill_only_rect_stays_behind_overlapping_text(self) -> None:
        block = IRBlock(
            kind="vector",
            page=1,
            bbox=(18.0, 18.0, 82.0, 32.0),
            z_order=10,
            vector=PdfVectorObject(
                kind="rect",
                bbox=(18.0, 18.0, 82.0, 32.0),
                fill_color=(255, 255, 255),
                closed=True,
                filled=True,
                stroked=False,
            ),
        )

        self.assertTrue(_filled_vector_is_behind_text(block, _page_with_text()))

    def test_grid_stroke_is_not_reclassified_as_text_background(self) -> None:
        block = IRBlock(
            kind="vector",
            page=1,
            bbox=(18.0, 24.0, 82.0, 25.0),
            z_order=10,
            vector=PdfVectorObject(
                kind="line",
                bbox=(18.0, 24.0, 82.0, 25.0),
                stroke_color=(0, 0, 0),
                stroke_width=0.75,
                stroked=True,
                filled=False,
            ),
        )

        self.assertFalse(_filled_vector_is_behind_text(block, _page_with_text()))

    def test_filled_path_with_late_z_order_remains_a_cover_candidate(self) -> None:
        block = IRBlock(
            kind="vector",
            page=1,
            bbox=(18.0, 18.0, 82.0, 32.0),
            z_order=10,
            vector=PdfVectorObject(
                kind="path",
                bbox=(18.0, 18.0, 82.0, 32.0),
                fill_color=(220, 220, 220),
                stroke_color=(0, 0, 0),
                closed=True,
                filled=True,
                stroked=True,
            ),
        )

        self.assertFalse(_filled_vector_is_behind_text(block, _page_with_text()))


if __name__ == "__main__":
    unittest.main()
