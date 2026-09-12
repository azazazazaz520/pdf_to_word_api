from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from src.validate.render import render_pdf_pages
from src.validate.similarity import compare_pdf_pages, ssim_score
from reportlab.pdfgen.canvas import Canvas

import numpy as np


class PageSimilarityTest(unittest.TestCase):
    def test_identical_images_score_one(self) -> None:
        image = np.zeros((120, 90), dtype=np.float32)
        image[20:60, 30:60] = 200.0
        self.assertAlmostEqual(ssim_score(image, image), 1.0, places=5)

    def test_noise_reduces_similarity(self) -> None:
        rng = np.random.default_rng(7)
        base = rng.integers(0, 255, size=(64, 64)).astype(np.float32)
        noisy = np.clip(base + rng.normal(0, 40, size=base.shape), 0, 255)
        score = ssim_score(base, noisy)
        self.assertLess(score, 0.9)
        self.assertGreater(score, 0.0)

    def test_different_pages_score_lower_than_identical_pdf(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_a = root / "a.pdf"
            pdf_b = root / "b.pdf"
            for path, text in ((pdf_a, "Alpha"), (pdf_b, "Beta")):
                canvas = Canvas(str(path), pagesize=(300, 400))
                canvas.setFont("Helvetica", 14)
                canvas.drawString(40, 340, text)
                canvas.save()

            same = compare_pdf_pages(pdf_a, pdf_a)
            different = compare_pdf_pages(pdf_a, pdf_b)

            self.assertEqual(same["status"], "succeeded")
            self.assertGreaterEqual(same["min_ssim"], 0.999)
            self.assertLess(different["min_ssim"], same["min_ssim"])
            self.assertEqual(different["compared_page_count"], 1)

    def test_render_pdf_pages_honours_page_selection(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            pdf_path = root / "two.pdf"
            canvas = Canvas(str(pdf_path), pagesize=(300, 400))
            canvas.drawString(40, 340, "first")
            canvas.showPage()
            canvas.drawString(40, 340, "second")
            canvas.save()

            images = render_pdf_pages(pdf_path, dpi=72.0, page_indices=[1])
            self.assertEqual(len(images), 1)
            self.assertGreater(images[0].shape[0], 0)


if __name__ == "__main__":
    unittest.main()
